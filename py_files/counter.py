import argparse
from datetime import datetime
from collections import defaultdict
import re
from pathlib import Path
from urllib.parse import urljoin

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

from playwright.sync_api import sync_playwright

# =========================================
# KONFIGURATION
# =========================================

FIRST_TEAM_URL = "https://www.fussball.de/mannschaft/sgm-mariazell-locherhof-stetten-lackendorf-sv-mariazell-wuerttemberg/-/saison/2526/team-id/02TCEJ3RA4000000VS5489BRVTHNGU03#!/"
MEN_TEAM_NAME_PREFIX = "SGM Mariazell"
SECOND_TEAM_URL = "https://www.fussball.de/mannschaft/sgm-mariazell-locherhof-stetten-lackendorf-ii-sv-mariazell-wuerttemberg/-/saison/2526/team-id/02TCEK4EA0000000VS5489BRVTHNGU03#!/"
WOMEN_TEAM_URL = "https://www.fussball.de/mannschaft/sgm-locherhof-mariazell-fv-locherhof-wuerttemberg/-/saison/2526/team-id/011MIB7LHO000000VTVG0001VTR8C1K7#!/"
WOMEN_TEAM_NAME_PREFIX = "SGM Locherhof"
DATE_FROM = "24.07.2025"
DATE_TO = "08.06.2026"
HEADLESS = False
WAIT_MS = 1200
DEBUG = False


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Laedt Fussball.de Daten und exportiert Spieler-Statistiken."
    )
    parser.add_argument(
        "--first-team-url",
        "--team-url",
        dest="first_team_url",
        default=FIRST_TEAM_URL,
        help="URL der 1. Mannschaft",
    )
    parser.add_argument(
        "--second-team-url",
        default=SECOND_TEAM_URL,
        help="URL der 2. Mannschaft (leer = ueberspringen)",
    )
    parser.add_argument(
        "--women-team-url",
        default=WOMEN_TEAM_URL,
        help="URL der Damen-Mannschaft (leer = ueberspringen)",
    )
    parser.add_argument(
        "--own-team-name-prefix",
        default=MEN_TEAM_NAME_PREFIX,
        help="Praefix fuer Teamzuordnung",
    )
    parser.add_argument(
        "--women-team-name-prefix",
        default=WOMEN_TEAM_NAME_PREFIX,
        help="Praefix fuer Damen-Teamzuordnung",
    )
    parser.add_argument("--date-from", default=DATE_FROM, help="Startdatum dd.mm.yyyy")
    parser.add_argument("--date-to", default=DATE_TO, help="Enddatum dd.mm.yyyy")
    parser.add_argument("--headless", action="store_true", help="Browser headless starten")
    parser.add_argument("--wait-ms", type=int, default=WAIT_MS, help="Wartezeit in ms")
    parser.add_argument("--debug", action="store_true", help="Debug-Ausgabe aktivieren")
    return parser.parse_args()


def normalize_text(value: str) -> str:
    value = value.lower()
    value = value.replace("ä", "ae")
    value = value.replace("ö", "oe")
    value = value.replace("ü", "ue")
    value = value.replace("ß", "ss")
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def get_own_team_keywords() -> list[str]:
    match = re.search(r"/mannschaft/([^/]+)/", FIRST_TEAM_URL)
    if not match:
        return []

    slug = match.group(1)
    parts = [normalize_text(part) for part in slug.split("-")]
    keywords = []
    for part in parts:
        if len(part) < 3:
            continue
        if part in {"sgm", "sv", "fv", "fc", "tsv", "sc", "spfr", "vfr", "tus", "sg"}:
            continue
        keywords.append(part)

    seen = set()
    unique_keywords = []
    for keyword in keywords:
        if keyword not in seen:
            seen.add(keyword)
            unique_keywords.append(keyword)
    return unique_keywords


def debug(msg: str) -> None:
    if DEBUG:
        print(f"[DEBUG] {msg}")


def accept_cookies_if_present(page) -> None:
    selectors = [
        "button:has-text('Akzeptiere alle')"
    ]
    for selector in selectors:
        try:
            loc = page.locator(selector).first
            if loc.is_visible(timeout=1200):
                loc.click()
                debug("Cookie-Banner akzeptiert")
                page.wait_for_timeout(600)
                return
        except Exception:
            pass


def click_first_visible(page, selectors, timeout=1200) -> bool:
    for selector in selectors:
        try:
            loc = page.locator(selector).first
            if loc.is_visible(timeout=timeout):
                loc.click()
                page.wait_for_timeout(800)
                return True
        except Exception:
            pass
    return False


def click_los_button(page) -> bool:
    # Datepicker ggf. schließen, damit der Button nicht überdeckt ist.
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass

    los_word_pattern = re.compile(r"\blos\b", re.IGNORECASE)

    # Nur echte Buttons pruefen und "Los" als eigenes Wort matchen (nicht z.B. "sieglos").
    try:
        candidates = page.locator("button")
        count = candidates.count()
        for i in range(count):
            el = candidates.nth(i)
            if not el.is_visible():
                continue

            try:
                if not el.is_enabled():
                    continue
            except Exception:
                pass

            text = (el.inner_text() or "").strip()
            value = (el.get_attribute("value") or "").strip()
            title = (el.get_attribute("title") or "").strip()
            aria = (el.get_attribute("aria-label") or "").strip()
            combined = " ".join(part for part in [text, value, title, aria] if part)

            if los_word_pattern.search(combined):
                try:
                    el.click(timeout=1000)
                except Exception:
                    # Fallback auf JS-Click.
                    el.evaluate("node => node.click()")
                return True
    except Exception:
        pass

    return False


def click_aufstellung_tab(page) -> bool:
    selector_candidates = [
        "a:has-text('Aufstellung')",
        "button:has-text('Aufstellung')",
        "[role='tab']:has-text('Aufstellung')",
        "a[title*='Aufstellung' i]",
        "button[title*='Aufstellung' i]",
    ]
    if click_first_visible(page, selector_candidates, timeout=2500):
        debug("Tab 'Aufstellung' geoeffnet")
        return True

    try:
        candidates = page.locator("a, button, [role='tab']")
        count = candidates.count()
        for i in range(count):
            el = candidates.nth(i)
            if not el.is_visible():
                continue
            text = (el.inner_text() or "").strip().lower()
            title = ((el.get_attribute("title") or "").strip()).lower()
            aria = ((el.get_attribute("aria-label") or "").strip()).lower()
            if "aufstellung" in text or "aufstellung" in title or "aufstellung" in aria:
                try:
                    el.click(timeout=1200)
                except Exception:
                    el.evaluate("node => node.click()")
                page.wait_for_timeout(1000)
                debug("Tab 'Aufstellung' geoeffnet")
                return True
    except Exception:
        pass

    debug("Tab 'Aufstellung' nicht gefunden")
    return False


def format_player_name(name: str) -> str:
    def cap_token(token: str) -> str:
        parts = token.split("-")
        capped = []
        for part in parts:
            if not part:
                capped.append(part)
            else:
                capped.append(part[0].upper() + part[1:].lower())
        return "-".join(capped)

    clean = re.sub(r"\s+", " ", (name or "").strip())
    if not clean:
        return clean
    return " ".join(cap_token(token) for token in clean.split(" "))


def format_player_name_for_excel(name: str) -> str:
    clean = format_player_name(name)
    if not clean or "," in clean:
        return clean

    parts = clean.split()
    if len(parts) < 2:
        return clean

    last_name = parts[-1]
    first_names = " ".join(parts[:-1])
    return f"{last_name}, {first_names}"


def player_sort_key(name: str) -> tuple[str, str]:
    formatted = format_player_name_for_excel(name)
    if "," in formatted:
        last, first = formatted.split(",", 1)
        return (normalize_text(last), normalize_text(first))
    return (normalize_text(formatted), "")


def period_token(value: str) -> str:
    key = normalize_date_key(value)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", key):
        return key[:4]

    match = re.search(r"(\d{4})", value or "")
    return match.group(1) if match else "unknown"


def extract_team_players(page, team_name_prefix: str) -> list[dict[str, str]]:
    lineup_data = page.evaluate(
        r"""
        ({ teamNamePrefix }) => {
            const normalize = (value) => {
                return (value || "")
                    .toLowerCase()
                    .normalize("NFD")
                    .replace(/[\u0300-\u036f]/g, "")
                    .replace(/[^a-z0-9]+/g, " ")
                    .replace(/\s+/g, " ")
                    .trim();
            };

            const targetNamePrefix = normalize(teamNamePrefix);
            const fieldWrapper = document.querySelector('.field-wrapper');
            const field = fieldWrapper ? fieldWrapper.querySelector('.field') : null;
            const homeSection = field ? field.querySelector('.home') : null;
            const awaySection = field ? field.querySelector('.away') : null;

            const readClubName = (section) => {
                if (!section) {
                    return '';
                }
                const club = section.querySelector('.club-name');
                return normalize(club ? (club.innerText || '') : '');
            };

            const homeClubName = readClubName(homeSection);
            const awayClubName = readClubName(awaySection);

            let ownSide = 'unbekannt';
            if (homeClubName.startsWith(targetNamePrefix)) {
                ownSide = 'home';
            } else if (awayClubName.startsWith(targetNamePrefix)) {
                ownSide = 'away';
            }

            const wrappers = Array.from(document.querySelectorAll('.player-wrapper.' + ownSide));
            const links = [];
            const readMinutesPlayed = (wrapper) => {
                const texts = [
                    wrapper.getAttribute('data-minute'),
                    wrapper.getAttribute('data-minutes'),
                    wrapper.getAttribute('title'),
                    wrapper.getAttribute('aria-label'),
                    wrapper.innerText,
                    wrapper.textContent,
                ];

                for (const raw of texts) {
                    if (!raw) {
                        continue;
                    }

                    const text = String(raw).replace(/\s+/g, ' ').trim();
                    const minuteMatch = text.match(/\b(\d{1,3}(?:\+\d{1,2})?)\s*(?:'|min\.?|minutes?)\b/i)
                        || text.match(/\b(\d{1,3}(?:\+\d{1,2})?)'/i);
                    if (minuteMatch) {
                        return minuteMatch[1];
                    }
                }

                return '';
            };

            const players = [];
            for (const wrapper of wrappers) {
                const anchor = wrapper.querySelector('a[href]');
                let href = '';
                if (anchor && anchor.href) {
                    href = anchor.href;
                } else {
                    const rawHref = wrapper.getAttribute('href');
                    if (rawHref) {
                        href = new URL(rawHref, window.location.href).href;
                    }
                }

                if (!href) {
                    continue;
                }

                links.push(href);
                players.push({
                    profileUrl: href,
                    minutesPlayed: readMinutesPlayed(wrapper),
                });
            }

            return {
                players,
                profileUrls: Array.from(new Set(links)),
            };
        }
        """,
        {"teamNamePrefix": team_name_prefix},
    )

    lineup_players = lineup_data.get("players", [])
    profile_links = lineup_data.get("profileUrls", [])
    if not profile_links:
        return []

    minutes_by_profile_url = {
        item.get("profileUrl"): item.get("minutesPlayed", "")
        for item in lineup_players
        if isinstance(item, dict) and item.get("profileUrl")
    }

    lineup_url = page.url
    seen_names = set()
    names = []

    for profile_url in profile_links:
        try:
            page.goto(profile_url, wait_until="domcontentloaded", timeout=120000)
            page.wait_for_timeout(900)
            name_loc = page.locator(".profile-name").first
            profile_name = (name_loc.inner_text(timeout=1200) or "").strip()
            if not profile_name:
                profile_name = (page.locator("h1").first.inner_text(timeout=1200) or "").strip()
            if not profile_name:
                continue

            profile_name = format_player_name(profile_name)
            key = profile_name.lower()
            if key in seen_names:
                continue
            seen_names.add(key)
            minutes_played = minutes_by_profile_url.get(profile_url, "")
            names.append({"name": profile_name, "minutes": minutes_played})
        except Exception:
            pass
        finally:
            try:
                page.goto(lineup_url, wait_until="domcontentloaded", timeout=120000)
                page.wait_for_timeout(600)
            except Exception:
                pass

    return names


def get_match_date(page, link: str) -> str:
    raw = page.evaluate(
        r"""
        () => {
            const linkSelectors = [
                '.stage.header a.competition',
                'a.competition[href*="/spieldatum/"]',
                'a[href*="/spieldatum/"]',
            ];

            for (const selector of linkSelectors) {
                const node = document.querySelector(selector);
                if (!node) {
                    continue;
                }
                const href = (node.getAttribute('href') || '').trim();
                const hrefIso = href.match(/spieldatum\/(\d{4}-\d{2}-\d{2})/i);
                if (hrefIso) {
                    return hrefIso[1];
                }
            }

            const selectors = [
                '.stage.header',
                '.match-stage',
                '.content',
                'body',
            ];
            for (const selector of selectors) {
                const node = document.querySelector(selector);
                if (!node) {
                    continue;
                }
                const text = (node.innerText || '').trim();
                const de = text.match(/\b\d{2}\.\d{2}\.\d{4}\b/);
                if (de) {
                    return de[0];
                }
                const iso = text.match(/\b\d{4}-\d{2}-\d{2}\b/);
                if (iso) {
                    return iso[0];
                }
            }
            return '';
        }
        """
    )

    raw = (raw or "").strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}$", raw):
        return raw

    if re.match(r"^\d{2}\.\d{2}\.\d{4}$", raw):
        dt = datetime.strptime(raw, "%d.%m.%Y")
        return dt.strftime("%Y-%m-%d")

    match = re.search(r"(\d{4}-\d{2}-\d{2})", link)
    if match:
        return match.group(1)

    return "unbekannt"


def normalize_date_key(value: str) -> str:
    text = (value or "").strip()
    if re.match(r"^\d{4}-\d{2}-\d{2}$", text):
        return text

    if re.match(r"^\d{2}\.\d{2}\.\d{4}$", text):
        dt = datetime.strptime(text, "%d.%m.%Y")
        return dt.strftime("%Y-%m-%d")

    embedded_iso = re.search(r"(\d{4}-\d{2}-\d{2})", text)
    if embedded_iso:
        return embedded_iso.group(1)

    return text.lower()


def format_date_for_output(value: str) -> str:
    key = normalize_date_key(value)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", key):
        dt = datetime.strptime(key, "%Y-%m-%d")
        return dt.strftime("%d.%m.%Y")
    return value


def get_first_visible_locator(page, selectors, timeout=1200):
    for selector in selectors:
        try:
            loc = page.locator(selector).first
            if loc.is_visible(timeout=timeout):
                return loc
        except Exception:
            pass
    return None


def set_readonly_date_value(page, selectors, value: str, label: str) -> bool:
    for selector in selectors:
        try:
            loc = page.locator(selector).first
            if loc.is_visible(timeout=1200):
                # Die Datumsfelder sind readonly, daher per JS setzen + Events ausloesen.
                page.eval_on_selector(
                    selector,
                    """
                    (el, val) => {
                        el.removeAttribute('readonly');
                        el.value = val;
                        el.dispatchEvent(new Event('input', { bubbles: true }));
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                        el.dispatchEvent(new Event('blur', { bubbles: true }));
                    }
                    """,
                    value,
                )
                debug(f"{label} gesetzt auf {value} via {selector}")
                return True
        except Exception:
            pass
    return False


def set_readonly_date_value_at_index(page, selector: str, index: int, value: str, label: str) -> bool:
    try:
        loc = page.locator(selector)
        if loc.count() <= index:
            return False

        target = loc.nth(index)
        if not target.is_visible(timeout=1200):
            return False

        target.evaluate(
            """
            (el, val) => {
                el.removeAttribute('readonly');
                el.value = val;
                el.dispatchEvent(new Event('input', { bubbles: true }));
                el.dispatchEvent(new Event('change', { bubbles: true }));
                el.dispatchEvent(new Event('blur', { bubbles: true }));
            }
            """,
            value,
        )
        debug(f"{label} gesetzt auf {value} via {selector}[{index}]")
        return True
    except Exception:
        return False


def parse_de_date(value: str):
    return datetime.strptime(value, "%d.%m.%Y")


def read_picker_month_year(page):
    title_selectors = [
        ".uib-datepicker-popup .uib-title strong",
        ".uib-datepicker-popup .uib-title",
        ".datepicker .datepicker-switch",
        ".dropdown-menu .uib-title",
    ]
    title_text = ""
    for sel in title_selectors:
        try:
            loc = page.locator(sel).first
            if loc.is_visible(timeout=500):
                title_text = (loc.inner_text() or "").strip()
                if title_text:
                    break
        except Exception:
            pass

    if not title_text:
        return None, None

    months = {
        "januar": 1,
        "februar": 2,
        "maerz": 3,
        "märz": 3,
        "april": 4,
        "mai": 5,
        "juni": 6,
        "juli": 7,
        "august": 8,
        "september": 9,
        "oktober": 10,
        "november": 11,
        "dezember": 12,
    }

    lower = title_text.lower()
    month = None
    for name, num in months.items():
        if name in lower:
            month = num
            break

    year = None
    for token in lower.replace(".", " ").split():
        if token.isdigit() and len(token) == 4:
            year = int(token)
            break

    return month, year


def navigate_picker_to_month(page, target_month: int, target_year: int) -> bool:
    prev_selectors = [
        ".uib-datepicker-popup .uib-left",
        ".datepicker .prev",
        "button[aria-label*='Vorheriger' i]",
        "button:has-text('‹')",
    ]
    next_selectors = [
        ".uib-datepicker-popup .uib-right",
        ".datepicker .next",
        "button[aria-label*='Naechster' i]",
        "button[aria-label*='Nächster' i]",
        "button:has-text('›')",
    ]

    for _ in range(240):
        current_month, current_year = read_picker_month_year(page)
        if current_month == target_month and current_year == target_year:
            return True
        if current_month is None or current_year is None:
            break

        current = (current_year, current_month)
        target = (target_year, target_month)
        if current < target:
            if not click_first_visible(page, next_selectors, timeout=700):
                return False
        else:
            if not click_first_visible(page, prev_selectors, timeout=700):
                return False
        page.wait_for_timeout(250)
    return False


def click_day_in_open_picker(page, day: int) -> bool:
    day_text = str(day)
    button_scopes = [
        ".uib-datepicker-popup button",
        ".datepicker button",
        ".dropdown-menu button",
    ]

    # Erst nicht-ausgegraute Tagesbuttons bevorzugen.
    for scope in button_scopes:
        try:
            buttons = page.locator(scope)
            count = buttons.count()
            for i in range(count):
                btn = buttons.nth(i)
                if not btn.is_visible():
                    continue
                txt = (btn.inner_text() or "").strip()
                if txt != day_text:
                    continue
                classes = (btn.get_attribute("class") or "").lower()
                if "muted" in classes or "disabled" in classes:
                    continue
                btn.click()
                return True
        except Exception:
            pass

    # Fallback: notfalls erster passender Tag.
    for scope in button_scopes:
        try:
            buttons = page.locator(scope)
            count = buttons.count()
            for i in range(count):
                btn = buttons.nth(i)
                if not btn.is_visible():
                    continue
                txt = (btn.inner_text() or "").strip()
                if txt == day_text:
                    btn.click()
                    return True
        except Exception:
            pass
    return False


def pick_date_via_calendar(page, selectors, value: str, label: str) -> bool:
    target = parse_de_date(value)
    field = get_first_visible_locator(page, selectors, timeout=1200)
    if field is None:
        return False

    try:
        field.click()
        page.wait_for_timeout(120)
    except Exception:
        return False

    moved = navigate_picker_to_month(page, target.month, target.year)
    if not moved:
        return False

    clicked = click_day_in_open_picker(page, target.day)
    if not clicked:
        return False

    page.wait_for_timeout(80)
    debug(f"{label} per Kalender gewaehlt: {value}")
    return True


def pick_date_via_calendar_at_index(page, selector: str, index: int, value: str, label: str) -> bool:
    try:
        loc = page.locator(selector)
        if loc.count() <= index:
            return False
        field = loc.nth(index)
        if not field.is_visible(timeout=1200):
            return False
        field.click()
        page.wait_for_timeout(120)
        target = parse_de_date(value)
        if not navigate_picker_to_month(page, target.month, target.year):
            return False
        if not click_day_in_open_picker(page, target.day):
            return False
        page.wait_for_timeout(80)
        debug(f"{label} per Kalender gewaehlt: {value} via {selector}[{index}]")
        return True
    except Exception:
        return False


def open_matchplan_set_dates_and_submit(page, team_url: str) -> None:
    debug("Oeffne Teamseite")
    page.goto(team_url, wait_until="domcontentloaded", timeout=120000)
    page.wait_for_timeout(1800)
    accept_cookies_if_present(page)

    opened = click_first_visible(
        page,
        [
            "text=Mannschaftsspielplan",
            "a:has-text('Mannschaftsspielplan')",
            "button:has-text('Mannschaftsspielplan')",
            "text=Spielplan",
        ],
    )
    debug(f"Mannschaftsspielplan geoeffnet: {'ja' if opened else 'nein'}")

    from_ok = pick_date_via_calendar(
        page,
        [
            "#matchplan-date-from",
            "input[id*='date-from']",
            "input[name='datum-von']",
            "input[data-ajaxmodel='datum-von']",
        ],
        DATE_FROM,
        "Startdatum",
    )
    if not from_ok:
        from_ok = set_readonly_date_value(
            page,
            [
                "#matchplan-date-from",
                "input[id*='date-from']",
                "input[name='datum-von']",
                "input[data-ajaxmodel='datum-von']",
            ],
            DATE_FROM,
            "Startdatum (Fallback JS)",
        )

    to_ok = pick_date_via_calendar(
        page,
        [
            "#matchplan-date-to",
            "input[id*='date-to']",
            "input[name='datum-bis']",
            "input[data-ajaxmodel='datum-bis']",
        ],
        DATE_TO,
        "Enddatum",
    )
    if not to_ok:
        to_ok = set_readonly_date_value(
            page,
            [
                "#matchplan-date-to",
                "input[id*='date-to']",
                "input[name='datum-bis']",
                "input[data-ajaxmodel='datum-bis']",
            ],
            DATE_TO,
            "Enddatum (Fallback JS)",
        )

    # Fallback: Manche Seiten haben fuer beide Datumsfelder aehnliche/gleiche Attribute.
    if not to_ok:
        to_ok = pick_date_via_calendar_at_index(
            page,
            "#matchplan-date-from",
            1,
            DATE_TO,
            "Enddatum (Fallback Kalender)",
        ) or pick_date_via_calendar_at_index(
            page,
            "input[name='datum-von']",
            1,
            DATE_TO,
            "Enddatum (Fallback Kalender)",
        )

    if not to_ok:
        to_ok = set_readonly_date_value_at_index(
            page,
            "#matchplan-date-from",
            1,
            DATE_TO,
            "Enddatum (Fallback JS)",
        ) or set_readonly_date_value_at_index(
            page,
            "input[name='datum-von']",
            1,
            DATE_TO,
            "Enddatum (Fallback JS)",
        )

    if not from_ok:
        debug("Startdatum-Feld nicht gefunden")
    if not to_ok:
        debug("Enddatum-Feld nicht gefunden")

    los_clicked = click_los_button(page)
    debug(f"Button 'Los' geklickt: {'ja' if los_clicked else 'nein'}")
    if not los_clicked:
        raise RuntimeError("Button 'Los' konnte nicht geklickt werden.")


def click_mehr_laden_until_done(page) -> int:
    selectors = [
        "button:has-text('Mehr Laden')",
        "button:has-text('Mehr laden')",
        "a:has-text('Mehr Laden')",
        "a:has-text('Mehr laden')",
        "button:has-text('Mehr anzeigen')",
        "a:has-text('Mehr anzeigen')",
    ]

    clicks = 0
    for _ in range(200):
        clicked = click_first_visible(page, selectors, timeout=1200)
        if not clicked:
            # Einmal scrollen und final pruefen, ob wirklich nichts mehr da ist.
            page.mouse.wheel(0, 4500)
            page.wait_for_timeout(700)
            if not click_first_visible(page, selectors, timeout=700):
                break
            clicks += 1
            page.wait_for_timeout(900)
            continue

        clicks += 1
        if clicks % 5 == 0:
            debug(f"Mehr-Laden Klicks: {clicks}")
        page.mouse.wheel(0, 3500)
        page.wait_for_timeout(800)

    return clicks


def collect_played_match_links(page):
    # Findet die ERSTE Tabelle direkt nach dem "Los" Button und sammelt ALLE Spiel-Links darin.
    # Die relevante Tabelle ist nicht immer die umschliessende Tabelle des Los-Buttons.
    # Daher wird die erste nachfolgende Tabelle gesucht, die tatsaechlich Spiel-Links enthaelt.
    
    links = []
    seen = set()
    
    # Finde den Los-Button und anschließend die erste Tabelle, extrahiere alle Links daraus
    link_list = page.evaluate(
        r"""
        () => {
            // Suche nach einem Element mit "Los" Text
            let losButton = null;
            const buttons = document.querySelectorAll('button, a, input[type="submit"], input[type="button"]');
            for (let b of buttons) {
                const label = (b.innerText || b.value || b.title || '').trim();
                if (/\blos\b/i.test(label)) {
                    losButton = b;
                    break;
                }
            }

            if (!losButton) {
                return [];
            }

            const tableHasMatchLinks = (table) => {
                if (!table) {
                    return false;
                }
                return !!table.querySelector('a[href*="/spiel/"], a[href*="/match/"]');
            };

            const tables = Array.from(document.querySelectorAll('table'));
            let table = null;

            for (const candidate of tables) {
                if (!tableHasMatchLinks(candidate)) {
                    continue;
                }

                const position = losButton.compareDocumentPosition(candidate);
                const isAfterLosButton = Boolean(position & Node.DOCUMENT_POSITION_FOLLOWING);
                if (isAfterLosButton) {
                    table = candidate;
                    break;
                }
            }

            if (!table) {
                table = tables.find(tableHasMatchLinks) || null;
            }

            if (!table) {
                return [];
            }

            // Extrahiere nur Links aus column-score, bei denen bereits ein Ergebnis steht.
            // Spiele ohne Ergebnis enthalten dort stattdessen meist keine Score-Spans.
            const links = [];
            const scoreAnchors = table.querySelectorAll(
                'td.column-score a[href*="/spiel/"], td.column-score a[href*="/match/"]'
            );
            for (let a of scoreAnchors) {
                const hasScoreLeft = !!a.querySelector('.score-left');
                const hasScoreRight = !!a.querySelector('.score-right');
                if (!hasScoreLeft || !hasScoreRight) {
                    continue;
                }

                const href = a.getAttribute('href');
                if (href) {
                    links.push(href);
                }
            }

            return links;
        }
        """
    )
    
    debug(f"Gefundene Links aus Tabelle nach Los-Button: {len(link_list)}")
    if link_list:
        for href in link_list:
            full = urljoin("https://www.fussball.de", href.strip())
            if full not in seen:
                seen.add(full)
                links.append(full)
                debug(f"  Link: ...{full[-80:]}")
    else:
        debug("Keine Tabelle nach Los-Button gefunden oder keine Links in Tabelle")

    debug(f"Gesamt eindeutige Spiel-Links: {len(links)}")
    return links


def get_match_type(page) -> str:
    return page.evaluate(
        """
        () => {
            const selectors = [
                '.stage.header a.competition',
                '.match-stage .stage.header a.competition',
                '.content .match-stage a.competition',
                'a.competition',
            ];

            let title = '';
            let href = '';

            for (const selector of selectors) {
                const node = document.querySelector(selector);
                if (!node) {
                    continue;
                }
                title = ((node.innerText || node.textContent || '')).toLowerCase().trim();
                href = ((node.getAttribute('href') || '')).toLowerCase().trim();
                if (title || href) {
                    break;
                }
            }

            if (!title) {
                const header = document.querySelector('.stage.header');
                title = ((header && (header.innerText || header.textContent)) || '').toLowerCase().trim();
            }

            const haystack = `${title} ${href}`;

            if (haystack.includes('freundschaftsspiel')) {
                return 'Freundschaftsspiel';
            }
            if (haystack.includes('pokal')) {
                return 'Pokalspiel';
            }
            if (haystack.includes('liga')) {
                return 'Punktspiel';
            }
            return 'Unbekannt';
        }
        """
    )


def collect_player_events_for_team(page, schedule_url: str, links, team_name_prefix: str, team_priority: int, team_label: str):
    events = []
    print(f"\n{team_label}: {len(links)} Spiele mit Ergebnis")

    for idx, link in enumerate(links, start=1):
        print(f"[{idx}/{len(links)}] Oeffne: {link}")
        page.goto(link, wait_until="domcontentloaded", timeout=120000)

        try:
            page.wait_for_selector("a.competition, .stage.header", timeout=5000)
        except Exception:
            pass

        match_type = get_match_type(page)
        match_date = get_match_date(page, link)
        print(f"Datum: {format_date_for_output(match_date)}")
        print(f"Spieltyp: {match_type}")

        opened_aufstellung = click_aufstellung_tab(page)
        page.wait_for_timeout(800)
        if not opened_aufstellung:
            page.goto(schedule_url, wait_until="domcontentloaded", timeout=120000)
            page.wait_for_timeout(800)
            continue

        players = extract_team_players(page, team_name_prefix)
        for player in players:
            events.append(
                {
                    "player": player["name"],
                    "date": match_date,
                    "type": match_type,
                    "team_priority": team_priority,
                    "minutes": player.get("minutes", ""),
                }
            )

        page.goto(schedule_url, wait_until="domcontentloaded", timeout=120000)
        page.wait_for_timeout(700)

    return events


def summarize_player_events(events) -> dict:
    per_player_per_day = defaultdict(dict)

    for event in events:
        player = event["player"]
        date = normalize_date_key(event["date"])
        current = per_player_per_day[player].get(date)

        if current is None or event["team_priority"] >= current["team_priority"]:
            per_player_per_day[player][date] = {
                "type": event["type"],
                "team_priority": event["team_priority"],
                "minutes": event.get("minutes", ""),
            }

    summary = {}
    for player, day_map in per_player_per_day.items():
        punktspiel = 0
        pokalspiel = 0
        freundschaftsspiel = 0

        for item in day_map.values():
            match_type = item["type"]
            if match_type == "Punktspiel":
                punktspiel += 1
            elif match_type == "Pokalspiel":
                pokalspiel += 1
            elif match_type == "Freundschaftsspiel":
                freundschaftsspiel += 1

        summary[player] = {
            "Punktspiel": punktspiel,
            "Pokalspiel": pokalspiel,
            "Freundschaftsspiel": freundschaftsspiel,
            "Gesamt": len(day_map),
        }

    return summary


def print_summary(summary: dict) -> None:
    print("\nSpieler-Statistik (1. und 2. Mannschaft kombiniert):")
    if not summary:
        print("Keine Daten gefunden.")
        return

    players = sorted(summary.keys())
    idx_width = max(3, len(str(len(players))))
    name_width = max(6, max(len(player) for player in players))
    total_width = max(6, max(len(str(summary[player]["Gesamt"])) for player in players))
    punkt_width = max(10, max(len(str(summary[player]["Punktspiel"])) for player in players))
    pokal_width = max(10, max(len(str(summary[player]["Pokalspiel"])) for player in players))
    freund_width = max(18, max(len(str(summary[player]["Freundschaftsspiel"])) for player in players))

    header = (
        f"{'Nr':>{idx_width}}  "
        f"{'Spieler':<{name_width}}  "
        f"{'Gesamt':>{total_width}}  "
        f"{'Punktspiel':>{punkt_width}}  "
        f"{'Pokalspiel':>{pokal_width}}  "
        f"{'Freundschaftsspiel':>{freund_width}}"
    )
    print(header)
    print("-" * len(header))

    for idx, player in enumerate(players, start=1):
        stats = summary[player]
        print(
            f"{idx:>{idx_width}}  "
            f"{player:<{name_width}}  "
            f"{stats['Gesamt']:>{total_width}}  "
            f"{stats['Punktspiel']:>{punkt_width}}  "
            f"{stats['Pokalspiel']:>{pokal_width}}  "
            f"{stats['Freundschaftsspiel']:>{freund_width}}"
        )


def export_results(summary: dict, all_events: list, output_prefix: str) -> Path:
    start_token = period_token(DATE_FROM)
    end_token = period_token(DATE_TO)
    season_folder = f"{start_token}_{end_token}"
    out_dir = Path(__file__).parent.parent / "player_stats" / "yearly_stats" / season_folder
    out_dir.mkdir(parents=True, exist_ok=True)
    out_file = out_dir / f"{output_prefix}_stats_{start_token}_{end_token}.xlsx"

    wb = openpyxl.Workbook()

    # ── Blatt 1: Zusammenfassung ──────────────────────────────────────
    ws_summary = wb.active
    ws_summary.title = "Zusammenfassung"

    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    headers = ["Nr", "Spieler", "Gesamt", "Punktspiel", "Pokalspiel", "Freundschaftsspiel"]
    for col, text in enumerate(headers, start=1):
        cell = ws_summary.cell(row=1, column=col, value=text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    players_sorted = sorted(summary.keys(), key=player_sort_key)
    for idx, player in enumerate(players_sorted, start=1):
        stats = summary[player]
        ws_summary.append([
            idx,
            format_player_name_for_excel(player),
            stats["Gesamt"],
            stats["Punktspiel"],
            stats["Pokalspiel"],
            stats["Freundschaftsspiel"],
        ])

    # Spaltenbreiten anpassen
    ws_summary.column_dimensions["A"].width = 6
    ws_summary.column_dimensions["B"].width = 30
    for col_letter in ("C", "D", "E", "F"):
        ws_summary.column_dimensions[col_letter].width = 18

    # ── Blatt 2: Rohdaten (Einzelereignisse) ─────────────────────────
    ws_events = wb.create_sheet(title="Einzelereignisse")
    event_headers = ["Spieler", "Datum", "Typ", "Minuten", "Mannschaft-Prioritaet"]
    for col, text in enumerate(event_headers, start=1):
        cell = ws_events.cell(row=1, column=col, value=text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    events_sorted = sorted(
        all_events,
        key=lambda event: (
            player_sort_key(event["player"]),
            normalize_date_key(event["date"]),
            event["type"],
        ),
    )
    for event in events_sorted:
        ws_events.append([
            format_player_name_for_excel(event["player"]),
            format_date_for_output(event["date"]),
            event["type"],
            event.get("minutes", ""),
            event["team_priority"],
        ])

    ws_events.column_dimensions["A"].width = 30
    ws_events.column_dimensions["B"].width = 14
    ws_events.column_dimensions["C"].width = 20
    ws_events.column_dimensions["D"].width = 12
    ws_events.column_dimensions["E"].width = 22

    wb.save(out_file)
    return out_file


def main() -> None:
    global FIRST_TEAM_URL, SECOND_TEAM_URL, WOMEN_TEAM_URL
    global MEN_TEAM_NAME_PREFIX, WOMEN_TEAM_NAME_PREFIX
    global DATE_FROM, DATE_TO, HEADLESS, WAIT_MS, DEBUG

    args = parse_args()
    FIRST_TEAM_URL = args.first_team_url
    SECOND_TEAM_URL = args.second_team_url
    WOMEN_TEAM_URL = args.women_team_url
    MEN_TEAM_NAME_PREFIX = args.own_team_name_prefix
    WOMEN_TEAM_NAME_PREFIX = args.women_team_name_prefix
    DATE_FROM = args.date_from
    DATE_TO = args.date_to
    HEADLESS = args.headless
    WAIT_MS = args.wait_ms
    DEBUG = args.debug

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=HEADLESS)
        page = browser.new_page()
        page.set_extra_http_headers(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/123.0 Safari/537.36"
                )
            }
        )

        herren_teams = [
            {
                "label": "1. Mannschaft",
                "url": FIRST_TEAM_URL,
                "prefix": MEN_TEAM_NAME_PREFIX,
                "priority": 1,
            }
        ]

        if SECOND_TEAM_URL.strip():
            herren_teams.append(
                {
                    "label": "2. Mannschaft",
                    "url": SECOND_TEAM_URL,
                    "prefix": MEN_TEAM_NAME_PREFIX,
                    "priority": 2,
                }
            )
        else:
            print("Hinweis: SECOND_TEAM_URL ist leer. 2. Mannschaft wird uebersprungen.")

        damen_events = []
        if WOMEN_TEAM_URL.strip():
            women_prefix = WOMEN_TEAM_NAME_PREFIX.strip() or MEN_TEAM_NAME_PREFIX
            women_team = {
                "label": "Damen",
                "url": WOMEN_TEAM_URL,
                "prefix": women_prefix,
                "priority": 1,
            }

            open_matchplan_set_dates_and_submit(page, women_team["url"])
            page.wait_for_timeout(WAIT_MS)
            schedule_url = page.url
            click_mehr_laden_until_done(page)
            links = collect_played_match_links(page)
            damen_events = collect_player_events_for_team(
                page,
                schedule_url,
                links,
                women_team["prefix"],
                women_team["priority"],
                women_team["label"],
            )
        else:
            print("Hinweis: WOMEN_TEAM_URL ist leer. Damen werden uebersprungen.")

        herren_events = []
        for team in herren_teams:
            open_matchplan_set_dates_and_submit(page, team["url"])
            page.wait_for_timeout(WAIT_MS)

            schedule_url = page.url
            click_mehr_laden_until_done(page)
            links = collect_played_match_links(page)

            team_events = collect_player_events_for_team(
                page,
                schedule_url,
                links,
                team["prefix"],
                team["priority"],
                team["label"],
            )
            herren_events.extend(team_events)

        if herren_events:
            print("\n=== Herren ===")
            herren_summary = summarize_player_events(herren_events)
            print_summary(herren_summary)
            herren_out_file = export_results(herren_summary, herren_events, output_prefix="herren")
            print(f"\nExport Herren gespeichert: {herren_out_file}")
        else:
            print("\nKeine Herren-Daten gefunden.")

        if damen_events:
            print("\n=== Damen ===")
            damen_summary = summarize_player_events(damen_events)
            print_summary(damen_summary)
            damen_out_file = export_results(damen_summary, damen_events, output_prefix="damen")
            print(f"\nExport Damen gespeichert: {damen_out_file}")
        else:
            print("\nKeine Damen-Daten gefunden.")

        browser.close()


if __name__ == "__main__":
    main()