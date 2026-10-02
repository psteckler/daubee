/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ 1092
(__unused_webpack_module, __webpack_exports__, __webpack_require__) {

/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   Nj: () => (/* binding */ detectLocaleDefaultIsoCode)
/* harmony export */ });
/* unused harmony exports resolveLocaleToIsoCode, buildLocaleCurrencyNumberFormat, buildCurrencyNumberFormat, replaceCurrencyBracket, resolveCurrencyFormat */
// Resolves a cell's own Currency-/Accounting-category number format string
// to the specific currency it displays, so annotateSelection (taskpane.ts)
// can restrict a unit annotation to match -- see TODO.md's "Restrict unit
// annotations on Currency- and Accounting-formatted cells" item for the
// full design this implements the "Format parsing" piece of.
//
// Excel encodes the chosen currency symbol as a "[$TOKEN-LOCALE]" bracket
// token somewhere in the format code -- not anchored to the start or end
// (confirmed against real Excel: choosing "$ French (Canada)" produces
// "#,##0.00 [$-fr-CA]", a suffix, not a prefix) -- where TOKEN is either a
// literal glyph ("$", "€", "£", ...) or a three-letter ISO 4217 code, and
// LOCALE is either a legacy hex LCID ("409") or a modern BCP-47 locale tag
// ("fr-CA"). A TOKEN that's already a 3-letter code needs no further
// lookup at all; a bare glyph needs LOCALE to say which currency it
// actually means, since more than one currency can share the same glyph.
//
// Two lookup tables (LCID -> ISO code, BCP-47 region -> ISO code) resolve
// LOCALE to an ISO 4217 code. REGION_TO_ISO_CODE covers every currency
// seeded in units.seed.yaml as of 2026-09-22 ("let's add all currencies
// supported by Excel"), one representative country/region per currency --
// compiled from general ISO 4217/ISO 3166-1 reference knowledge, not
// verified cell-by-cell against Excel's own currency-format dropdown, so
// treat an unexpected miss or a wrong-seeming region as a data bug to
// report/fix rather than a fundamental limitation (the mechanism itself is
// sound; a specific mapping might just be off). LCID_TO_ISO_CODE, by
// contrast, is deliberately left at its original small size: it's already
// UNVERIFIED against real Excel (Paul's own confirmed example used the
// modern BCP-47 form, not a legacy hex LCID), and modern Excel's own
// output appears to prefer BCP-47 anyway, so expanding a table this
// speculative to match REGION_TO_ISO_CODE's own new breadth didn't seem
// worth the added risk of wrong hex values -- add a specific LCID if
// Paul's own testing actually hits one.
//
// A format with no bracket at all, but that still contains a bare or
// quoted currency-like glyph (e.g. `"$"#,##0.00`, Excel's own simplest
// Currency format for a user's own default locale), is structurally
// ambiguous -- which specific dollar-family currency does a bare "$"
// mean? -- and no amount of seeding more currencies would ever resolve
// that. Deliberately NOT treated as resolvable, even when only one
// seeded currency happens to use that glyph today: that would make the
// answer depend on which currencies happen to be registered, rather than
// on what the format text actually says.
// Legacy hex LCID -> ISO 4217 code, keyed without leading zeros (Excel's
// own bracket text -- e.g. "[$-409]" -- doesn't pad them). Only the
// primary English-language LCID for each seeded country is listed here;
// another language's LCID for the same country (e.g. French Canada,
// "C0C") isn't yet covered -- add it if Paul's own testing hits it.
const LCID_TO_ISO_CODE = {
    "409": "USD", // en-US
    "1009": "CAD", // en-CA
    C09: "AUD", // en-AU
    "1409": "NZD", // en-NZ
    "40C": "EUR", // fr-FR
    "407": "EUR", // de-DE
    "40D": "ILS", // he-IL
};
// BCP-47 region subtag -> ISO 4217 code -- the region alone (the part
// after the last "-") is enough, regardless of which language tag
// precedes it ("fr-CA" and "en-CA" both mean CAD). One representative
// region per seeded currency (Eurozone countries all point at EUR, the two
// CFA franc zones at their own respective codes, etc.) -- a currency
// shared by several countries doesn't need every one of them listed, only
// enough to resolve LOCALE's region subtag to the right code; add a
// specific country if a gap actually comes up.
const REGION_TO_ISO_CODE = {
    US: "USD",
    CA: "CAD",
    AU: "AUD",
    NZ: "NZD",
    FR: "EUR",
    DE: "EUR",
    ES: "EUR",
    IT: "EUR",
    IE: "EUR",
    PT: "EUR",
    NL: "EUR",
    AT: "EUR",
    BE: "EUR",
    FI: "EUR",
    GR: "EUR",
    AD: "EUR",
    IL: "ILS",
    AE: "AED",
    AF: "AFN",
    AG: "XCD",
    AL: "ALL",
    AM: "AMD",
    AO: "AOA",
    AR: "ARS",
    AW: "AWG",
    AZ: "AZN",
    BA: "BAM",
    BB: "BBD",
    BD: "BDT",
    BG: "BGN",
    BH: "BHD",
    BI: "BIF",
    BM: "BMD",
    BN: "BND",
    BO: "BOB",
    BR: "BRL",
    BS: "BSD",
    BT: "BTN",
    BW: "BWP",
    BY: "BYN",
    BZ: "BZD",
    CD: "CDF",
    CH: "CHF",
    CL: "CLP",
    CM: "XAF",
    CN: "CNY",
    CO: "COP",
    CR: "CRC",
    CU: "CUP",
    CV: "CVE",
    CW: "ANG",
    CZ: "CZK",
    DJ: "DJF",
    DK: "DKK",
    DO: "DOP",
    DZ: "DZD",
    EG: "EGP",
    ER: "ERN",
    ET: "ETB",
    FJ: "FJD",
    FK: "FKP",
    GB: "GBP",
    GE: "GEL",
    GH: "GHS",
    GI: "GIP",
    GM: "GMD",
    GN: "GNF",
    GT: "GTQ",
    GY: "GYD",
    HK: "HKD",
    HN: "HNL",
    HR: "HRK",
    HT: "HTG",
    HU: "HUF",
    ID: "IDR",
    IN: "INR",
    IQ: "IQD",
    IR: "IRR",
    IS: "ISK",
    JM: "JMD",
    JO: "JOD",
    JP: "JPY",
    KE: "KES",
    KG: "KGS",
    KH: "KHR",
    KM: "KMF",
    KP: "KPW",
    KR: "KRW",
    KW: "KWD",
    KY: "KYD",
    KZ: "KZT",
    LA: "LAK",
    LB: "LBP",
    LK: "LKR",
    LR: "LRD",
    LS: "LSL",
    LY: "LYD",
    MA: "MAD",
    MD: "MDL",
    MG: "MGA",
    MK: "MKD",
    MM: "MMK",
    MN: "MNT",
    MO: "MOP",
    MR: "MRU",
    MU: "MUR",
    MV: "MVR",
    MW: "MWK",
    MX: "MXN",
    MY: "MYR",
    MZ: "MZN",
    NA: "NAD",
    NG: "NGN",
    NI: "NIO",
    NO: "NOK",
    NP: "NPR",
    OM: "OMR",
    PA: "PAB",
    PE: "PEN",
    PF: "XPF",
    PG: "PGK",
    PH: "PHP",
    PK: "PKR",
    PL: "PLN",
    PY: "PYG",
    QA: "QAR",
    RO: "RON",
    RS: "RSD",
    RU: "RUB",
    RW: "RWF",
    SA: "SAR",
    SB: "SBD",
    SC: "SCR",
    SD: "SDG",
    SE: "SEK",
    SG: "SGD",
    SH: "SHP",
    SL: "SLE",
    SN: "XOF",
    SO: "SOS",
    SR: "SRD",
    SS: "SSP",
    ST: "STN",
    SV: "SVC",
    SY: "SYP",
    SZ: "SZL",
    TH: "THB",
    TJ: "TJS",
    TM: "TMT",
    TN: "TND",
    TO: "TOP",
    TR: "TRY",
    TT: "TTD",
    TW: "TWD",
    TZ: "TZS",
    UA: "UAH",
    UG: "UGX",
    UY: "UYU",
    UZ: "UZS",
    VE: "VES",
    VN: "VND",
    VU: "VUV",
    WS: "WST",
    YE: "YER",
    ZA: "ZAR",
    ZM: "ZMW",
};
// Bare/quoted currency glyphs to watch for when no "[$...]" bracket is
// present at all -- see the class-level comment above for why a match
// here always means "ambiguous," never a resolved ISO code. Covers every
// single-character, non-letter glyph any seeded currency uses (a
// letter-based abbreviation like Swedish krona's "kr" or Polish zloty's
// "zł" is deliberately excluded -- matching those as a bare substring
// anywhere in a format string would false-positive on all sorts of
// unrelated text, unlike a real currency glyph).
const BARE_CURRENCY_GLYPHS = (/* unused pure expression or super */ null && ([
    "$",
    "€",
    "₪",
    "£",
    "¥",
    "₽",
    "₴",
    "₺",
    "₹",
    "₨",
    "৳",
    "₩",
    "฿",
    "₫",
    "₱",
    "៛",
    "₭",
    "₮",
    "₸",
    "₼",
    "₾",
    "֏",
    "؋",
    "﷼",
    "₦",
    "₵",
    "₣",
    "₲",
    "₡",
    "ƒ",
]));
// Matches Excel's own "[$TOKEN-LOCALE]" bracket syntax anywhere in a
// format string. TOKEN (group 1, possibly empty -- "[$-409]" omits it,
// meaning "this locale's own default symbol") is everything up to the
// first "-" inside the brackets; LOCALE (group 2, optional -- absent
// when the bracket has no "-" at all, e.g. a bare "[$USD]") is
// everything after that "-". TOKEN itself never contains "-", so
// splitting on the first one is safe.
const CURRENCY_BRACKET_PATTERN = /\[\$([^\]-]*)(?:-([^\]]*))?\]/;
// Extracts the region subtag from a BCP-47-ish locale tag by taking
// everything after the last "-" (works for "fr-CA" -> "CA", and equally
// for a bare region with no language tag at all, e.g. "CA" -> "CA").
function regionSubtag(locale) {
    const parts = locale.split("-");
    return parts[parts.length - 1];
}
function resolveLocaleToIsoCode(locale) {
    const trimmed = locale.trim();
    if (/^[0-9A-Fa-f]+$/.test(trimmed)) {
        // Looks like a legacy hex LCID (digits and A-F only, no letters that
        // couldn't also be hex) -- normalize case/leading zeros before
        // matching the table's own keys.
        const normalized = trimmed.replace(/^0+(?=.)/, "").toUpperCase();
        const byLcid = LCID_TO_ISO_CODE[normalized];
        if (byLcid) {
            return byLcid;
        }
    }
    const region = regionSubtag(trimmed).toUpperCase();
    return REGION_TO_ISO_CODE[region];
}
// Best-effort locale-based currency guess for the SECURITIES_UNIT-requiring
// securities-pricing function family's own default (see PreferenceStore's
// securitiesUnitIsoCode and resolveSecuritiesUnit, taskpane.ts) and for the
// Settings dialog's own "Locale default (...)" display -- both call this
// so neither can ever disagree about what "locale default" currently
// means. navigator.language is a plain browser API, safe to call from
// any page including a dialog's own restricted Office.js context (unlike
// Office.context.document); undefined here just means "couldn't guess,"
// callers fall back to a hardcoded default (US_dollar) the same way they
// always did before this preference existed.
function detectLocaleDefaultIsoCode() {
    const locale = typeof navigator !== "undefined" ? navigator.language : undefined;
    return locale ? resolveLocaleToIsoCode(locale) : undefined;
}
// Symbol + one representative BCP-47 locale for each seeded currency, used
// by buildLocaleCurrencyNumberFormat below to write a "[$SYMBOL-LOCALE]"
// bracket that displays the currency's own familiar glyph (e.g. "$42.99")
// rather than a bare ISO code ("USD 42.99") -- Paul's own preferred
// default (2026-09-22), after live-testing confirmed buildCurrencyNumberFormat's
// bare-ISO-code bracket really does render the literal 3-letter code as
// text rather than a mapped symbol, the "users are used to seeing '$
// 42.99', not 'USD 42.99'" issue he flagged. Confirmed live example for
// USD: "[$$-en-US]#,##0.00". Only one representative locale is picked per
// currency (e.g. "en-US" for USD, arbitrarily "fr-FR" over "de-DE" for
// EUR) -- where several locales share a currency, the choice has no effect
// on which digits/symbol a plain numeric Currency format like this shows.
// Covers every currency seeded in units.seed.yaml as of 2026-09-22 -- same
// general-reference-knowledge, not-cell-by-cell-verified caveat as
// REGION_TO_ISO_CODE above applies to the symbol/locale choices here too.
//
// CONFIRMED RISK (2026-09-24): a bare "$" was wrong for HKD -- real Excel's
// own Currency dropdown (Format Cells > Currency > "Chinese (Traditional,
// Hong Kong SAR)") produces "[$HK$-zh-HK]#,##0.00", i.e. the symbol itself
// is "HK$", not "$" (this app's own resolveCurrencyFormat still resolves a
// bare "$" guess correctly via the locale's own region tag when READING an
// existing format, so the functional damage is limited to what gets
// WRITTEN by the no-format-yet/rewrite-token offers -- but Excel's own
// Format Cells dialog then mislabels it, e.g. as some unrelated locale that
// happens to share the plain "$" glyph).
//
// A same-day batch of "well-established real-world disambiguating symbol"
// guesses for seven more currencies turned out to be a genuinely unreliable
// heuristic, not just an unverified one -- Paul checked every one of them
// against his own Excel, and the FINAL tally is 3 confirmed correct (TWD
// "NT$", HKD "HK$", XCD "EC$" -- each a real, matching dropdown entry)
// against 5 confirmed WRONG (SGD, JMD, TTD -- real raw formats
// "[$$-en-SG]#,##0.00"/"[$$-en-JM]#,##0.00"/"[$$-en-TT]#,##0.00", a
// matching dropdown entry found directly each time; NAD, FJD -- no
// matching dropdown entry at all for "N$"/"FJ$"), all five since reverted
// back to bare "$". Worse than a coin flip. Paul's own verdict once the
// batch was fully checked: "I think this approach makes sense. If a user
// wants to distinguish the currencies from other dollars, they can always
// use the ISO abbreviations for formatting" (see the useIsoCodeCurrencyFormat
// preference, preferences.ts) -- i.e. this heuristic is retired for good;
// no more proactively guessing a disambiguating symbol from general
// knowledge. Every bare-"$" entry below (the large majority of the dollar-
// and peso-family currencies in this table, including Barbados, Bermuda,
// Cayman Islands, Liberia, Solomon Islands, Suriname, Cape Verde, and the
// whole Latin American peso group) should be treated as genuinely unknown,
// not "probably fine" -- fix one in place only once Paul has actually
// checked it against real Excel, the same way every currency above was --
// see TODO.md's own currency item for the tracked list.
const ISO_CODE_TO_SYMBOL_LOCALE = (/* unused pure expression or super */ null && ({
    USD: { symbol: "$", locale: "en-US" },
    CAD: { symbol: "$", locale: "en-CA" },
    AUD: { symbol: "$", locale: "en-AU" },
    NZD: { symbol: "$", locale: "en-NZ" },
    EUR: { symbol: "€", locale: "fr-FR" },
    ILS: { symbol: "₪", locale: "he-IL" },
    GBP: { symbol: "£", locale: "en-GB" },
    JPY: { symbol: "¥", locale: "ja-JP" },
    CHF: { symbol: "Fr", locale: "de-CH" },
    CNY: { symbol: "¥", locale: "zh-CN" },
    HKD: { symbol: "HK$", locale: "zh-HK" },
    SGD: { symbol: "$", locale: "en-SG" },
    SEK: { symbol: "kr", locale: "sv-SE" },
    NOK: { symbol: "kr", locale: "nb-NO" },
    DKK: { symbol: "kr", locale: "da-DK" },
    PLN: { symbol: "zł", locale: "pl-PL" },
    CZK: { symbol: "Kč", locale: "cs-CZ" },
    HUF: { symbol: "Ft", locale: "hu-HU" },
    RON: { symbol: "lei", locale: "ro-RO" },
    BGN: { symbol: "лв", locale: "bg-BG" },
    HRK: { symbol: "kn", locale: "hr-HR" },
    ISK: { symbol: "kr", locale: "is-IS" },
    RUB: { symbol: "₽", locale: "ru-RU" },
    UAH: { symbol: "₴", locale: "uk-UA" },
    TRY: { symbol: "₺", locale: "tr-TR" },
    INR: { symbol: "₹", locale: "en-IN" },
    PKR: { symbol: "₨", locale: "ur-PK" },
    BDT: { symbol: "৳", locale: "bn-BD" },
    LKR: { symbol: "₨", locale: "si-LK" },
    NPR: { symbol: "₨", locale: "ne-NP" },
    KRW: { symbol: "₩", locale: "ko-KR" },
    TWD: { symbol: "NT$", locale: "zh-TW" },
    THB: { symbol: "฿", locale: "th-TH" },
    VND: { symbol: "₫", locale: "vi-VN" },
    IDR: { symbol: "Rp", locale: "id-ID" },
    MYR: { symbol: "RM", locale: "ms-MY" },
    PHP: { symbol: "₱", locale: "en-PH" },
    MMK: { symbol: "K", locale: "my-MM" },
    KHR: { symbol: "៛", locale: "km-KH" },
    LAK: { symbol: "₭", locale: "lo-LA" },
    MNT: { symbol: "₮", locale: "mn-MN" },
    KZT: { symbol: "₸", locale: "kk-KZ" },
    UZS: { symbol: "so'm", locale: "uz-UZ" },
    AZN: { symbol: "₼", locale: "az-AZ" },
    GEL: { symbol: "₾", locale: "ka-GE" },
    AMD: { symbol: "֏", locale: "hy-AM" },
    AFN: { symbol: "؋", locale: "fa-AF" },
    IRR: { symbol: "﷼", locale: "fa-IR" },
    IQD: { symbol: "ع.د", locale: "ar-IQ" },
    SAR: { symbol: "﷼", locale: "ar-SA" },
    AED: { symbol: "د.إ", locale: "ar-AE" },
    QAR: { symbol: "﷼", locale: "ar-QA" },
    KWD: { symbol: "د.ك", locale: "ar-KW" },
    BHD: { symbol: ".د.ب", locale: "ar-BH" },
    OMR: { symbol: "﷼", locale: "ar-OM" },
    JOD: { symbol: "د.ا", locale: "ar-JO" },
    LBP: { symbol: "ل.ل", locale: "ar-LB" },
    SYP: { symbol: "£", locale: "ar-SY" },
    YER: { symbol: "﷼", locale: "ar-YE" },
    EGP: { symbol: "£", locale: "ar-EG" },
    MAD: { symbol: "د.م.", locale: "ar-MA" },
    DZD: { symbol: "د.ج", locale: "ar-DZ" },
    TND: { symbol: "د.ت", locale: "ar-TN" },
    LYD: { symbol: "ل.د", locale: "ar-LY" },
    ZAR: { symbol: "R", locale: "en-ZA" },
    NGN: { symbol: "₦", locale: "en-NG" },
    GHS: { symbol: "₵", locale: "en-GH" },
    KES: { symbol: "KSh", locale: "en-KE" },
    TZS: { symbol: "TSh", locale: "sw-TZ" },
    UGX: { symbol: "USh", locale: "en-UG" },
    ETB: { symbol: "Br", locale: "am-ET" },
    XOF: { symbol: "CFA", locale: "fr-SN" },
    XAF: { symbol: "FCFA", locale: "fr-CM" },
    XPF: { symbol: "₣", locale: "fr-PF" },
    ZMW: { symbol: "ZK", locale: "en-ZM" },
    MZN: { symbol: "MT", locale: "pt-MZ" },
    AOA: { symbol: "Kz", locale: "pt-AO" },
    BWP: { symbol: "P", locale: "en-BW" },
    NAD: { symbol: "$", locale: "en-NA" },
    MUR: { symbol: "₨", locale: "en-MU" },
    MWK: { symbol: "MK", locale: "en-MW" },
    RWF: { symbol: "FRw", locale: "rw-RW" },
    XCD: { symbol: "EC$", locale: "en-AG" },
    BRL: { symbol: "R$", locale: "pt-BR" },
    MXN: { symbol: "$", locale: "es-MX" },
    ARS: { symbol: "$", locale: "es-AR" },
    CLP: { symbol: "$", locale: "es-CL" },
    COP: { symbol: "$", locale: "es-CO" },
    PEN: { symbol: "S/", locale: "es-PE" },
    UYU: { symbol: "$", locale: "es-UY" },
    BOB: { symbol: "Bs", locale: "es-BO" },
    PYG: { symbol: "₲", locale: "es-PY" },
    VES: { symbol: "Bs", locale: "es-VE" },
    GTQ: { symbol: "Q", locale: "es-GT" },
    CRC: { symbol: "₡", locale: "es-CR" },
    DOP: { symbol: "RD$", locale: "es-DO" },
    JMD: { symbol: "$", locale: "en-JM" },
    TTD: { symbol: "$", locale: "en-TT" },
    BSD: { symbol: "$", locale: "en-BS" },
    BZD: { symbol: "BZ$", locale: "en-BZ" },
    PAB: { symbol: "B/.", locale: "es-PA" },
    HNL: { symbol: "L", locale: "es-HN" },
    NIO: { symbol: "C$", locale: "es-NI" },
    SVC: { symbol: "$", locale: "es-SV" },
    CUP: { symbol: "$", locale: "es-CU" },
    HTG: { symbol: "G", locale: "fr-HT" },
    FJD: { symbol: "$", locale: "en-FJ" },
    PGK: { symbol: "K", locale: "en-PG" },
    WST: { symbol: "T", locale: "en-WS" },
    TOP: { symbol: "T$", locale: "to-TO" },
    BAM: { symbol: "KM", locale: "bs-BA" },
    MKD: { symbol: "ден", locale: "mk-MK" },
    RSD: { symbol: "дин", locale: "sr-RS" },
    ALL: { symbol: "L", locale: "sq-AL" },
    MDL: { symbol: "L", locale: "ro-MD" },
    BYN: { symbol: "Br", locale: "be-BY" },
    GIP: { symbol: "£", locale: "en-GI" },
    CDF: { symbol: "FC", locale: "fr-CD" },
    SDG: { symbol: "ج.س.", locale: "ar-SD" },
    SSP: { symbol: "£", locale: "en-SS" },
    SOS: { symbol: "Sh", locale: "so-SO" },
    DJF: { symbol: "Fdj", locale: "fr-DJ" },
    ERN: { symbol: "Nfk", locale: "ti-ER" },
    LSL: { symbol: "L", locale: "en-LS" },
    SZL: { symbol: "L", locale: "en-SZ" },
    SCR: { symbol: "₨", locale: "en-SC" },
    STN: { symbol: "Db", locale: "pt-ST" },
    CVE: { symbol: "$", locale: "pt-CV" },
    GMD: { symbol: "D", locale: "en-GM" },
    GNF: { symbol: "FG", locale: "fr-GN" },
    LRD: { symbol: "$", locale: "en-LR" },
    SLE: { symbol: "Le", locale: "en-SL" },
    BIF: { symbol: "FBu", locale: "fr-BI" },
    KMF: { symbol: "CF", locale: "fr-KM" },
    MGA: { symbol: "Ar", locale: "mg-MG" },
    MVR: { symbol: "Rf", locale: "dv-MV" },
    BTN: { symbol: "Nu.", locale: "dz-BT" },
    BND: { symbol: "$", locale: "ms-BN" },
    MOP: { symbol: "MOP$", locale: "zh-MO" },
    KPW: { symbol: "₩", locale: "ko-KP" },
    ANG: { symbol: "ƒ", locale: "nl-CW" },
    AWG: { symbol: "ƒ", locale: "nl-AW" },
    BBD: { symbol: "$", locale: "en-BB" },
    BMD: { symbol: "$", locale: "en-BM" },
    FKP: { symbol: "£", locale: "en-FK" },
    GYD: { symbol: "$", locale: "en-GY" },
    KGS: { symbol: "с", locale: "ky-KG" },
    KYD: { symbol: "$", locale: "en-KY" },
    MRU: { symbol: "UM", locale: "ar-MR" },
    SBD: { symbol: "$", locale: "en-SB" },
    SHP: { symbol: "£", locale: "en-SH" },
    SRD: { symbol: "$", locale: "nl-SR" },
    TJS: { symbol: "SM", locale: "tg-TJ" },
    TMT: { symbol: "m", locale: "tk-TM" },
    VUV: { symbol: "VT", locale: "en-VU" },
}));
// Builds just the "[$TOKEN-LOCALE]" bracket (no surrounding number
// pattern) for isoCode -- the currency's own symbol+locale
// (ISO_CODE_TO_SYMBOL_LOCALE) unless preferIsoCode is true or no such
// entry exists yet, in which case the bare 3-letter ISO code is used
// instead. Shared by buildLocaleCurrencyNumberFormat/
// buildCurrencyNumberFormat below (a full format string, for the
// "no format chosen yet" offer) and replaceCurrencyBracket further down
// (swapping just the bracket within an existing format string, for the
// rewrite-mismatched-format offer).
function buildCurrencyBracket(isoCode, preferIsoCode) {
    if (!preferIsoCode) {
        const entry = ISO_CODE_TO_SYMBOL_LOCALE[isoCode];
        if (entry) {
            return `[$${entry.symbol}-${entry.locale}]`;
        }
    }
    return `[$${isoCode}]`;
}
// Builds a Currency-style number format string using the currency's own
// familiar symbol plus a representative locale tag (e.g.
// "[$$-en-US]#,##0.00" for USD) -- the preferred default for the "no
// format chosen yet" offer (maybeApplyCurrencyFormat, taskpane.ts; see
// TODO.md's "No-format-yet case" sub-item), confirmed live to render as
// "$42.99" rather than "USD 42.99". Returns undefined for an ISO code with
// no ISO_CODE_TO_SYMBOL_LOCALE entry yet -- callers fall back to
// buildCurrencyNumberFormat's bare ISO-code form in that case.
function buildLocaleCurrencyNumberFormat(isoCode) {
    if (!ISO_CODE_TO_SYMBOL_LOCALE[isoCode]) {
        return undefined;
    }
    return `${buildCurrencyBracket(isoCode, false)}#,##0.00`;
}
// Builds a Currency-style number format string using the bare 3-letter ISO
// 4217 code (e.g. "[$USD]#,##0.00") rather than the currency's own symbol
// -- the same shape resolveCurrencyFormat's own ISO-code branch below
// already resolves without needing a locale, so a cell formatted this way
// round-trips correctly back through resolveCurrencyFormat on a later
// annotation. Confirmed live (2026-09-22) to render as the literal 3-letter
// code ("USD 42.99"), not a mapped symbol -- which is why
// buildLocaleCurrencyNumberFormat above is the preferred default; this is
// now the explicit opt-in (the useIsoCodeCurrencyFormat preference,
// preferences.ts) and the fallback for a currency with no symbol/locale
// entry yet.
function buildCurrencyNumberFormat(isoCode) {
    return `${buildCurrencyBracket(isoCode, true)}#,##0.00`;
}
// How many "[$TOKEN-LOCALE]" brackets format contains.
function countCurrencyBrackets(format) {
    const matches = format.match(new RegExp(CURRENCY_BRACKET_PATTERN.source, "g"));
    return matches ? matches.length : 0;
}
// Swaps the single currency bracket in format for one built for isoCode
// (buildCurrencyBracket), leaving everything else in the format string --
// decimals, thousands separator, negative/color sections, Accounting's own
// padding/alignment characters -- untouched. The write side of the
// mismatched-currency-format rewrite offer (annotateSelection,
// taskpane.ts; see TODO.md's own "offer ... to update the cell's number
// format's TOKEN" sub-item) -- rewriting both TOKEN and LOCALE together,
// rather than just TOKEN as that item's own first draft assumed, since two
// currencies can share the same glyph (CAD and AUD both use "$"): leaving
// the old LOCALE in place after only swapping the symbol would still
// resolve back to the *original* currency, not the new one.
//
// Returns undefined -- meaning "don't offer a rewrite, this format is too
// unusual to touch with confidence" -- whenever format doesn't contain
// exactly one bracket: zero (nothing to swap) or more than one (e.g. a
// custom negative-section color that repeats the bracket) both bail,
// rather than risk a partial or wrong rewrite.
function replaceCurrencyBracket(format, isoCode, preferIsoCode) {
    if (countCurrencyBrackets(format) !== 1) {
        return undefined;
    }
    const match = CURRENCY_BRACKET_PATTERN.exec(format);
    if (!match) {
        return undefined;
    }
    const bracket = buildCurrencyBracket(isoCode, preferIsoCode);
    return (format.slice(0, match.index) +
        bracket +
        format.slice(match.index + match[0].length));
}
function resolveCurrencyFormat(format) {
    const bracketMatch = CURRENCY_BRACKET_PATTERN.exec(format);
    if (bracketMatch) {
        const token = bracketMatch[1];
        const locale = bracketMatch[2];
        if (/^[A-Za-z]{3}$/.test(token)) {
            return { kind: "resolved", isoCode: token.toUpperCase() };
        }
        if (locale) {
            const isoCode = resolveLocaleToIsoCode(locale);
            if (isoCode) {
                return { kind: "resolved", isoCode };
            }
        }
        return { kind: "ambiguous" };
    }
    if (BARE_CURRENCY_GLYPHS.some((glyph) => format.includes(glyph))) {
        return { kind: "ambiguous" };
    }
    return { kind: "none" };
}


/***/ }

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	const __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		const cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		const module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter/value functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			if(Array.isArray(definition)) {
/******/ 				var i = 0;
/******/ 				while(i < definition.length) {
/******/ 					var key = definition[i++];
/******/ 					var binding = definition[i++];
/******/ 					if(!__webpack_require__.o(exports, key)) {
/******/ 						if(binding === 0) {
/******/ 							Object.defineProperty(exports, key, { enumerable: true, value: definition[i++] });
/******/ 						} else {
/******/ 							Object.defineProperty(exports, key, { enumerable: true, get: binding });
/******/ 						}
/******/ 					} else if(binding === 0) { i++; }
/******/ 				}
/******/ 			} else {
/******/ 				for(var key in definition) {
/******/ 					if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 						Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 					}
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/************************************************************************/
let __webpack_exports__ = {};
/* harmony import */ var _taskpane_currencyFormat__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(1092);
// Standalone page opened via Office.context.ui.displayDialogAsync from
// commands.ts's own "openSettings" ExecuteFunction handler (the ribbon's
// "Settings" button). Dialog windows run in a separate, restricted Office.js
// runtime with no Office.context.document at all -- confirmed the hard way
// (PreferenceStore.load() threw "Cannot read properties of undefined
// (reading 'settings')" when called directly from here); only
// Office.context.ui (messageParent/addHandlerAsync) is available in a
// dialog's own context. So this page can't touch PreferenceStore itself --
// it messages commands.ts (which does have full Office.context.document
// access, and stays loaded as the add-in's FunctionFile for the whole Excel
// session) to read/write on its behalf, the same ready-ping
// request/response shape notice-dialog.ts uses with its own parent
// (taskpane.ts there, commands.ts here).

// Cleared once the actual preference values arrive -- see the "ready" ping
// below.
let readyIntervalId;
Office.onReady(() => {
    const autoApplyCheckbox = document.getElementById("auto-apply-currency-format");
    autoApplyCheckbox.onchange = () => {
        Office.context.ui.messageParent(JSON.stringify({
            type: "set",
            key: "autoApplyCurrencyFormat",
            value: autoApplyCheckbox.checked,
        }));
    };
    const useIsoCodeCheckbox = document.getElementById("use-iso-code-currency-format");
    useIsoCodeCheckbox.onchange = () => {
        Office.context.ui.messageParent(JSON.stringify({
            type: "set",
            key: "useIsoCodeCurrencyFormat",
            value: useIsoCodeCheckbox.checked,
        }));
    };
    const saveUnitAliasesCheckbox = document.getElementById("save-unit-aliases");
    saveUnitAliasesCheckbox.onchange = () => {
        Office.context.ui.messageParent(JSON.stringify({
            type: "set",
            key: "saveUnitAliases",
            value: saveUnitAliasesCheckbox.checked,
        }));
    };
    const simplifyCheckbox = document.getElementById("simplify-using-unit-definitions");
    simplifyCheckbox.onchange = () => {
        Office.context.ui.messageParent(JSON.stringify({
            type: "set",
            key: "simplifyUsingUnitDefinitions",
            value: simplifyCheckbox.checked,
        }));
    };
    const autoFormatCheckbox = document.getElementById("auto-format-with-units");
    autoFormatCheckbox.onchange = () => {
        Office.context.ui.messageParent(JSON.stringify({
            type: "set",
            key: "autoFormatWithUnits",
            value: autoFormatCheckbox.checked,
        }));
    };
    const unitFormatStyleSelect = document.getElementById("unit-format-style");
    unitFormatStyleSelect.onchange = () => {
        Office.context.ui.messageParent(JSON.stringify({
            type: "set",
            key: "unitFormatStyle",
            value: unitFormatStyleSelect.value,
        }));
    };
    const unitFormatPrecisionSelect = document.getElementById("unit-format-precision");
    unitFormatPrecisionSelect.onchange = () => {
        Office.context.ui.messageParent(JSON.stringify({
            type: "set",
            key: "unitFormatPrecision",
            value: Number(unitFormatPrecisionSelect.value),
        }));
    };
    const securitiesUnitPinCheckbox = document.getElementById("securities-unit-pin-currency");
    const securitiesUnitSelect = document.getElementById("securities-unit-iso-code");
    // Shared by both controls: the checkbox picks "any currency" (unchecked)
    // vs. "a specific one" (checked); the select, only meaningful/enabled
    // when checked, then picks *which* -- "Locale default" (value "") or an
    // explicit currency. The select's own value is read regardless of the
    // checkbox's state (not reset when unchecking) so a previously-picked
    // currency survives a round trip through "any" and back, in case the
    // user re-checks the box -- resolveSecuritiesUnit (taskpane.ts) only ever
    // reads securitiesUnitIsoCode when the mode is actually "fixed", so storing
    // it while the mode is "any" is harmless.
    const sendSecuritiesUnitPreference = () => {
        securitiesUnitSelect.disabled = !securitiesUnitPinCheckbox.checked;
        const mode = !securitiesUnitPinCheckbox.checked
            ? "any"
            : securitiesUnitSelect.value === ""
                ? "locale"
                : "fixed";
        // One combined message, not two separate "set" ones -- commands.ts's
        // own PreferenceStore.setSecuritiesUnit writes both fields in a single
        // read-merge-write round trip. Two independent "set" messages raced
        // (each re-reads fresh state right before its own write, so the
        // second could read before the first's write had landed and silently
        // revert it) -- confirmed live: picking "Any currency" reverted to
        // "Locale default" instead of sticking.
        Office.context.ui.messageParent(JSON.stringify({
            type: "setSecuritiesUnit",
            mode,
            isoCode: securitiesUnitSelect.value,
        }));
    };
    securitiesUnitPinCheckbox.onchange = sendSecuritiesUnitPreference;
    securitiesUnitSelect.onchange = sendSecuritiesUnitPreference;
    document.getElementById("settings-close").onclick =
        () => {
            Office.context.ui.messageParent(JSON.stringify({ type: "closed" }));
        };
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, onParentMessage);
    // commands.ts's own handler only starts listening once displayDialogAsync's
    // callback fires, which can happen after this page has already loaded and
    // pinged once -- so keep pinging until the actual values arrive, rather
    // than assuming a single ping landed.
    const sendReady = () => Office.context.ui.messageParent(JSON.stringify({ type: "ready" }));
    sendReady();
    readyIntervalId = window.setInterval(sendReady, 200);
});
// Rebuilds the currency dropdown's own options from scratch every time
// preferences arrive (the currency list can change between opens --
// enabling/disabling a Unit Dictionaries group, or adding a custom
// dictionary -- so nothing here is cached across calls). Two kinds of
// option: "Locale default" (labeled with whatever navigator.language
// actually resolves to right now, via detectLocaleDefaultIsoCode,
// currencyFormat.ts, so it never claims a currency
// configureSecuritiesUnitFromPreferences (taskpane.ts) wouldn't also pick --
// both call the exact same function), then every other registered
// currency alphabetically. "Any currency" (see PreferenceStore's own
// securitiesUnitMode doc comment) isn't one of the select's own options --
// it's the checkbox being unchecked instead (see
// sendSecuritiesUnitPreference above), which also disables this select
// entirely, since which currency it's set to doesn't matter in that mode.
function renderSecuritiesUnitOptions(currencies, mode, selectedIsoCode) {
    const select = document.getElementById("securities-unit-iso-code");
    select.innerHTML = "";
    const detectedIsoCode = (0,_taskpane_currencyFormat__WEBPACK_IMPORTED_MODULE_0__/* .detectLocaleDefaultIsoCode */ .Nj)();
    const detected = detectedIsoCode
        ? currencies.find((c) => c.isoCode === detectedIsoCode)
        : undefined;
    const localeOption = document.createElement("option");
    localeOption.value = "";
    localeOption.textContent = detected
        ? `Locale default (${detected.name}, ${detected.isoCode})`
        : "Locale default (falls back to US dollar)";
    select.appendChild(localeOption);
    // The detected currency is already represented by the "Locale default"
    // option above -- listing it again as its own ordinary entry would let
    // a user pick the same effective currency two different ways, one of
    // which (the plain entry) wouldn't actually track the locale if it
    // ever changed.
    for (const currency of currencies) {
        if (currency.isoCode === detectedIsoCode) {
            continue;
        }
        const option = document.createElement("option");
        option.value = currency.isoCode;
        option.textContent = `${currency.name} (${currency.isoCode})`;
        select.appendChild(option);
    }
    // selectedIsoCode is shown regardless of mode (not just "fixed") so a
    // previously-picked currency is still visible, just disabled, after
    // closing and reopening the dialog while in "any" mode -- matching
    // sendSecuritiesUnitPreference's own within-session behavior of never
    // clearing it just because the checkbox is unchecked.
    select.value = selectedIsoCode;
    document.getElementById("securities-unit-pin-currency").checked = mode !== "any";
    select.disabled = mode === "any";
}
function onParentMessage(arg) {
    const data = JSON.parse(arg.message);
    if (data.type === "preferences") {
        if (readyIntervalId !== undefined) {
            window.clearInterval(readyIntervalId);
            readyIntervalId = undefined;
        }
        document.getElementById("auto-apply-currency-format").checked = data.autoApplyCurrencyFormat;
        document.getElementById("use-iso-code-currency-format").checked = data.useIsoCodeCurrencyFormat;
        document.getElementById("save-unit-aliases").checked =
            data.saveUnitAliases;
        document.getElementById("simplify-using-unit-definitions").checked = data.simplifyUsingUnitDefinitions;
        document.getElementById("auto-format-with-units").checked = data.autoFormatWithUnits;
        document.getElementById("unit-format-style").value =
            data.unitFormatStyle;
        document.getElementById("unit-format-precision").value = String(data.unitFormatPrecision);
        renderSecuritiesUnitOptions(data.currencies, data.securitiesUnitMode, data.securitiesUnitIsoCode);
    }
}

/******/ })()
;