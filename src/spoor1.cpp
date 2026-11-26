/*
CASE E - SPOOR 1 (C++-versie): het cyclisch rooster met namen (afdelingen A t/m D)
=================================================================================

Dit programma doet exact hetzelfde als spoor1.py: dezelfde input, dezelfde stappen,
dezelfde output (byte voor byte dezelfde txt-bestanden).

IDEE
  Case D zegt hoeveel lijnen er nodig zijn: voor afdeling A 52 lijnen = 52 FTE.
  Een lijn is een standaardweek: een shift (Early, Late of Night) op 5 vaste dagen.
  Dit programma vult die lijnen in:
    1. de eigen verpleegkundigen krijgen elk een lijn (zo goed mogelijk passend bij
       hun cyclische voorkeur);
    2. het tekort in FTE (52 - wat het eigen personeel samen werkt) wordt aangeworven.
  Omdat deeltijdsen minder werken dan hun lijn, zijn er meer mensen dan lijnen.

INPUT  (uit de Excel van de groep, in data/)
  - 'Case D fix (2)'                 : de lijnen (3 x 9h, 5-2 non-identical), de werkpatronen,
                                       de starturen van de shiften en de minimale bezetting
  - 'Department X - Nurse Char.'     : per verpleegkundige: type, tewerkstelling, cyclische voorkeur
  - 'Department X - Preferences'     : per verpleegkundige: aversie per dag en shift (28 dagen)
  C++ kan zelf geen Excel lezen. Een .xlsx is een zip-bestand met XML-bestanden erin, dus dit
  programma pakt de xlsx zelf uit (DEEL A onderaan). Er zijn geen extra bibliotheken nodig.

PLAATS IN HET PROJECT
  project/
    data/Excel_Group11_Case_A_B_C_D.xlsx   <- input (wordt hier gelezen)
    src/spoor1.cpp                         <- dit bestand (naast main.cpp)
    files/                                 <- output: invoer voor roster.exe
    results/                               <- output: overzichten
  Het programma zoekt zelf de projectmap (de map met 'data/' erin): eerst vanaf de map waar je
  het start, daarna vanaf de map van spoor1.exe.

OUTPUT
  - map files/ met de 4 invoerbestanden voor de C++-template:
      Shift_system_dpt_X.txt   shiften (starturen, lengte) en minimale bezetting per shift
      Cyclic_roster_dpt_X.txt  het cyclisch rooster: per persoon 28 dagen (standaardweek x 4)
      Personnel_dpt_X.txt      per persoon: nummer, 28 x 5 voorkeurscores, tewerkstelling, type
      Constraints_dpt_X.txt    roosterregels (min/max diensten, opeenvolgende dagen, ...)
  - map results/ met:
      spoor1_overzicht_dpt_X.xlsx              leesbaar overzicht van het resultaat
      Cyclic_roster_dpt_X_leesbaar.txt         hetzelfde cyclisch rooster, maar met kolomkoppen
                                               (dagen), rijkoppen (personen) en de bezetting per dag
  (Het bestand in files/ blijft zonder koppen: roster.exe leest het met fscanf en verwacht
   exact dat formaat.)

COMPILEREN  (apart van main.cpp: elk programma heeft zijn eigen main())
  Visual Studio, in de "Developer Command Prompt", vanuit de projectmap:
      cl /EHsc /std:c++17 /O2 src\spoor1.cpp /Fe:spoor1.exe
  g++ / MinGW (versie 9 of nieuwer), vanuit de projectmap:
      g++ -std=c++17 -O2 src/spoor1.cpp -o spoor1.exe

GEBRUIK  (vanuit de projectmap)
  spoor1.exe                                 -> alle afdelingen (DEPARTMENTS)
  spoor1.exe --dept A                        -> een afdeling
  spoor1.exe --dept A --excel pad\naar\bestand.xlsx
  (ook: --out <map voor de templatebestanden>, --results <map voor de overzichten>)

OPBOUW VAN DIT BESTAND  (dezelfde stappen en functienamen als spoor1.py)
  INSTELLINGEN
  DEEL A  - de xlsx lezen (unzip + XML) en de xlsx van het overzicht schrijven
  STAP 0  - inlezen            read_patterns, read_shifts, read_lines, read_nurses, ...
  STAP 1  - strafscore         parse_preference, score
  STAP 2  - eigen personeel    assign_own
  STAP 3  - aanwerven          contracts, fte, hire
  STAP 4  - uitvoer            write_template_files, write_overview, write_readable_roster
  HOOFDPROGRAMMA               run_department, project_root, main
*/

#define _CRT_SECURE_NO_WARNINGS

#include <algorithm>
#include <cctype>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <functional>
#include <iostream>
#include <iterator>
#include <map>
#include <optional>
#include <regex>
#include <set>
#include <sstream>
#include <stdexcept>
#include <string>
#include <tuple>
#include <unordered_map>
#include <utility>
#include <vector>

namespace fs = std::filesystem;
using std::string;
using std::vector;

// ============================================================================
// INSTELLINGEN
// ============================================================================

const string EXCEL_FILE = "Excel_Group11_Case_A_B_C_D.xlsx";   // bestandsnaam van de Excel in data/
const string DEPARTMENTS = "ABCD";                             // afdelingen die doorgerekend worden
const string DATA_DIR = "data";                                // map met de Excel
const string FILES_DIR = "files";                              // map waar roster.exe zijn input leest
const string RESULTS_DIR = "results";                          // map voor de overzichten

const string SHEET_CASE_D = "Case D fix (2)";
const string SCENARIO = "3 x 9h";
const string POLICY = "5-2 non-identical";
const int WEEKS = 4;                  // de standaardweek wordt 4 keer herhaald (28 dagen)

// Strafpunten voor de cyclische voorkeur. Contractbepalingen ("No ...", "Only ...") zijn hard.
const int PEN_SHIFT_PREF = 5;         // "Preference for X shifts" en de lijn heeft een ander shifttype
const int PEN_DAY_OFF = 1;            // per dag van de lijn die ze liever vrij heeft
const int PEN_DAY_ON = 1;             // per gewenste werkdag die niet in de lijn zit
const int PEN_TYPE1_ON_TYPE2 = 3;     // type 1 op een type II-lijn: mag, maar liever niet

const vector<double> HIRE_CONTRACTS = {0.5, 0.75, 1.0};   // contracten voor aanwervingen
const int HIRE_PREF = 5;                                   // voorkeurscore van een aanwerving

const bool REQ_WITH_BUFFER = false;   // Shift_system-bestand: minimum Case C (false) of eis met buffer (true)

struct ConstraintSettings {           // Constraints-bestand; de template vermenigvuldigt min/max_ass met de tewerkstelling
    int min_ass = 20, max_ass = 20;               // doel: 20 diensten per 28 dagen voor een voltijdse
    int min_cons_wrk = 1, max_cons_wrk = 6;       // opeenvolgende werkdagen
    int min_cons_shift = 1, max_cons_shift = 6;   // opeenvolgende dagen met dezelfde shift
    int min_shift = 0, max_shift = 28;            // aantal keer dezelfde shift per maand
    string identical = "N";                       // non-identical weekends
};
const ConstraintSettings CONSTRAINTS{};

const vector<string> DAYS_NL = {"Ma", "Di", "Wo", "Do", "Vr", "Za", "Zo"};
const vector<string> DAY_NAMES = {"monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"};
const vector<int> WEEKEND = {5, 6};
const std::map<string, string> LETTER = {{"Early", "E"}, {"Day", "D"}, {"Late", "L"}, {"Night", "N"}};
const std::regex ID_RE("30[12][A-D]\\d{3}");
const double EPS = 1e-9;


// ============================================================================
// HULPFUNCTIES (zodat de uitvoer exact gelijk is aan die van Python)
// ============================================================================

struct ExitError : std::runtime_error {            // = SystemExit in Python
    using std::runtime_error::runtime_error;
};

// Python repr(float): kortste schrijfwijze die exact hetzelfde getal teruggeeft (1.0, 0.75, 27.5, ...)
string py_repr(double x) {
    if (std::isnan(x)) return "nan";
    if (std::isinf(x)) return x > 0 ? "inf" : "-inf";
    char buf[64];
    for (int prec = 1; prec <= 17; ++prec) {
        std::snprintf(buf, sizeof buf, "%.*e", prec - 1, x);
        if (std::strtod(buf, nullptr) == x) break;
    }
    string s = buf;                                 // bv. "-2.75e+01"
    bool neg = s[0] == '-';
    if (neg) s = s.substr(1);
    size_t e = s.find('e');
    int exp = std::atoi(s.c_str() + e + 1);
    string digits;
    for (size_t i = 0; i < e; ++i)
        if (s[i] != '.') digits += s[i];
    while (digits.size() > 1 && digits.back() == '0') digits.pop_back();
    string out;
    if (exp >= -4 && exp < 16) {
        if (exp >= 0) {
            string ip = digits.substr(0, std::min<size_t>(digits.size(), exp + 1));
            while ((int)ip.size() < exp + 1) ip += '0';
            string fp = (int)digits.size() > exp + 1 ? digits.substr(exp + 1) : "0";
            out = ip + "." + fp;
        } else {
            out = "0." + string(-exp - 1, '0') + digits;
        }
    } else {
        out = digits.substr(0, 1);
        if (digits.size() > 1) out += "." + digits.substr(1);
        char eb[16];
        std::snprintf(eb, sizeof eb, "e%c%02d", exp < 0 ? '-' : '+', std::abs(exp));
        out += eb;
    }
    return (neg ? "-" : "") + out;
}

size_t u8len(const string& s) {                     // aantal tekens (zoals len() in Python)
    size_t n = 0;
    for (unsigned char c : s)
        if ((c & 0xC0) != 0x80) ++n;
    return n;
}

string pad_left(const string& s, size_t w) {        // f"{s:>w}"
    size_t n = u8len(s);
    return n >= w ? s : string(w - n, ' ') + s;
}

string pad_right(const string& s, size_t w) {       // f"{s:<w}"
    size_t n = u8len(s);
    return n >= w ? s : s + string(w - n, ' ');
}

string strip(const string& s) {                     // str.strip(): witruimte (ook niet-ASCII) weg
    static const vector<string> ws = {" ", "\t", "\n", "\r", "\x0b", "\x0c", "\x1c", "\x1d", "\x1e", "\x1f",
                                      "\xc2\x85", "\xc2\xa0", "\xe1\x9a\x80", "\xe2\x80\xa8", "\xe2\x80\xa9",
                                      "\xe2\x80\xaf", "\xe2\x81\x9f", "\xe3\x80\x80",
                                      "\xe2\x80\x80", "\xe2\x80\x81", "\xe2\x80\x82", "\xe2\x80\x83",
                                      "\xe2\x80\x84", "\xe2\x80\x85", "\xe2\x80\x86", "\xe2\x80\x87",
                                      "\xe2\x80\x88", "\xe2\x80\x89", "\xe2\x80\x8a"};
    size_t a = 0, b = s.size();
    bool changed = true;
    while (changed && a < b) {
        changed = false;
        for (const string& w : ws)
            if (b - a >= w.size() && s.compare(a, w.size(), w) == 0) { a += w.size(); changed = true; break; }
    }
    changed = true;
    while (changed && a < b) {
        changed = false;
        for (const string& w : ws)
            if (b - a >= w.size() && s.compare(b - w.size(), w.size(), w) == 0) { b -= w.size(); changed = true; break; }
    }
    return s.substr(a, b - a);
}

string lower(string s) {
    for (char& c : s)
        if (c >= 'A' && c <= 'Z') c = char(c - 'A' + 'a');
    return s;
}

string capitalize(const string& s) {               // "early" -> "Early"
    string t = lower(s);
    if (!t.empty() && t[0] >= 'a' && t[0] <= 'z') t[0] = char(t[0] - 'a' + 'A');
    return t;
}

bool starts_with(const string& s, const string& p) { return s.compare(0, p.size(), p) == 0; }
bool contains(const string& s, const string& p) { return s.find(p) != string::npos; }

string join(const vector<string>& v, const string& sep) {
    string out;
    for (size_t i = 0; i < v.size(); ++i) out += (i ? sep : "") + v[i];
    return out;
}

long long py_int_from_string(const string& raw, const string& what) {   // int("12")
    string s = strip(raw);
    size_t i = 0;
    if (i < s.size() && (s[i] == '+' || s[i] == '-')) ++i;
    if (i == s.size()) throw std::runtime_error("ongeldige gehele waarde '" + raw + "' (" + what + ")");
    for (size_t k = i; k < s.size(); ++k)
        if (s[k] < '0' || s[k] > '9') throw std::runtime_error("ongeldige gehele waarde '" + raw + "' (" + what + ")");
    return std::stoll(s);
}


// Een celwaarde, zoals openpyxl ze teruggeeft: None, int, float, str, bool of een tijdstip
struct Value {
    enum Kind { NONE, INT, FLOAT, STR, BOOL, TIME } kind = NONE;
    long long i = 0;      // INT, BOOL (0/1), TIME (milliseconden sinds middernacht)
    double f = 0;         // FLOAT
    string s;             // STR

    bool is_str() const { return kind == STR; }
    bool is_number() const { return kind == INT || kind == FLOAT || kind == BOOL; }   // isinstance(v, (int, float))
    bool operator==(const string& t) const { return kind == STR && s == t; }
    bool operator!=(const string& t) const { return !(*this == t); }
    bool truthy() const {
        switch (kind) {
            case NONE: return false;
            case INT: case BOOL: return i != 0;
            case FLOAT: return f != 0.0;
            case STR: return !s.empty();
            default: return true;
        }
    }
    string str() const {                               // str(v)
        switch (kind) {
            case NONE: return "None";
            case INT: return std::to_string(i);
            case FLOAT: return py_repr(f);
            case STR: return s;
            case BOOL: return i ? "True" : "False";
            default: {
                char b[16];
                long long sec = i / 1000;
                std::snprintf(b, sizeof b, "%02lld:%02lld:%02lld", sec / 3600 % 24, sec / 60 % 60, sec % 60);
                return b;
            }
        }
    }
    int hour() const { return int(i / 3600000 % 24); }
};

long long py_int(const Value& v, const string& what) {        // int(v)
    switch (v.kind) {
        case Value::INT: case Value::BOOL: return v.i;
        case Value::FLOAT:
            if (!std::isfinite(v.f)) throw std::runtime_error("ongeldig getal (" + what + ")");
            return (long long)std::trunc(v.f);
        case Value::STR: return py_int_from_string(v.s, what);
        default: throw std::runtime_error("geen getal in cel (" + what + "): " + v.str());
    }
}

double py_float(const Value& v, const string& what) {         // float(v)
    switch (v.kind) {
        case Value::INT: case Value::BOOL: return double(v.i);
        case Value::FLOAT: return v.f;
        case Value::STR: {
            string s = strip(v.s);
            char* end = nullptr;
            double d = std::strtod(s.c_str(), &end);
            if (s.empty() || *end != '\0') throw std::runtime_error("ongeldig getal '" + v.s + "' (" + what + ")");
            return d;
        }
        default: throw std::runtime_error("geen getal in cel (" + what + "): " + v.str());
    }
}

long long py_round(const Value& v, const string& what) {      // round(v): .5 naar het even getal
    switch (v.kind) {
        case Value::INT: case Value::BOOL: return v.i;
        case Value::FLOAT: return (long long)std::nearbyint(v.f);
        default: throw std::runtime_error("geen getal in cel (" + what + "): " + v.str());
    }
}

// Een getal of tekst in de samenvatting/het overzicht, met dezelfde str() als Python
struct PyVal {
    enum Kind { INT, FLOAT, STR } kind = STR;
    long long i = 0;
    double f = 0;
    string s;
    static PyVal I(long long x) { PyVal v; v.kind = INT; v.i = x; return v; }
    static PyVal F(double x) { PyVal v; v.kind = FLOAT; v.f = x; return v; }
    static PyVal S(const string& x) { PyVal v; v.kind = STR; v.s = x; return v; }
    string str() const { return kind == INT ? std::to_string(i) : kind == FLOAT ? py_repr(f) : s; }
    bool truthy() const { return kind == INT ? i != 0 : kind == FLOAT ? f != 0.0 : !s.empty(); }
    double num() const { return kind == INT ? double(i) : f; }
};

PyVal py_sum(const vector<double>& xs) {     // sum([...]) van tewerkstellingen: 0 (int) als de lijst leeg is
    if (xs.empty()) return PyVal::I(0);
    double t = 0;
    for (double x : xs) t += x;
    return PyVal::F(t);
}

PyVal py_round2(const PyVal& v) {            // round(x, 2)
    if (v.kind != PyVal::FLOAT) return v;
    char b[64];
    std::snprintf(b, sizeof b, "%.2f", v.f);
    return PyVal::F(std::strtod(b, nullptr));
}


// ============================================================================
// DEEL A - DE XLSX LEZEN EN SCHRIJVEN
//   Een .xlsx is een zip-archief met XML-bestanden. Om het te lezen:
//   (1) zip openen en de nodige bestanden uitpakken (inflate = deflate-decompressie),
//   (2) de XML lezen: workbook.xml (welke tabbladen), sharedStrings.xml (de teksten),
//       styles.xml (welke cellen een tijdformaat hebben) en de tabbladen zelf.
//   Dit deel moet je niet begrijpen om de case te volgen; het doet wat openpyxl doet in Python.
// ============================================================================

uint32_t crc32(const uint8_t* p, size_t n) {
    static uint32_t table[256];
    static bool init = false;
    if (!init) {
        for (uint32_t k = 0; k < 256; ++k) {
            uint32_t c = k;
            for (int j = 0; j < 8; ++j) c = (c & 1) ? 0xEDB88320u ^ (c >> 1) : c >> 1;
            table[k] = c;
        }
        init = true;
    }
    uint32_t c = 0xFFFFFFFFu;
    for (size_t k = 0; k < n; ++k) c = table[(c ^ p[k]) & 0xFF] ^ (c >> 8);
    return c ^ 0xFFFFFFFFu;
}

// Deflate-decompressie (RFC 1951), naar het voorbeeld van 'puff' van Mark Adler
class Inflater {
public:
    Inflater(const uint8_t* in, size_t n, vector<uint8_t>& out) : in_(in), n_(n), out_(out) {}

    void run() {
        int last;
        do {
            last = bits(1);
            int type = bits(2);
            if (type == 0) stored();
            else if (type == 1) fixed();
            else if (type == 2) dynamic();
            else throw std::runtime_error("zip: ongeldig blok");
        } while (!last);
    }

private:
    struct Huffman {
        short count[16] = {0};
        vector<short> symbol;
    };

    const uint8_t* in_;
    size_t n_, pos_ = 0;
    uint32_t bitbuf_ = 0;
    int bitcnt_ = 0;
    vector<uint8_t>& out_;

    int bits(int need) {
        uint32_t val = bitbuf_;
        while (bitcnt_ < need) {
            if (pos_ >= n_) throw std::runtime_error("zip: gegevens afgekapt");
            val |= uint32_t(in_[pos_++]) << bitcnt_;
            bitcnt_ += 8;
        }
        bitbuf_ = val >> need;
        bitcnt_ -= need;
        return int(val & ((1u << need) - 1));
    }

    static int construct(Huffman& h, const short* length, int n) {
        h.symbol.assign(n, 0);
        for (int len = 0; len < 16; ++len) h.count[len] = 0;
        for (int s = 0; s < n; ++s) h.count[length[s]]++;
        if (h.count[0] == n) return 0;
        int left = 1;
        for (int len = 1; len < 16; ++len) {
            left <<= 1;
            left -= h.count[len];
            if (left < 0) return left;
        }
        short offs[16];
        offs[1] = 0;
        for (int len = 1; len < 15; ++len) offs[len + 1] = short(offs[len] + h.count[len]);
        for (int s = 0; s < n; ++s)
            if (length[s] != 0) h.symbol[offs[length[s]]++] = short(s);
        return left;
    }

    int decode(const Huffman& h) {
        int code = 0, first = 0, index = 0;
        for (int len = 1; len < 16; ++len) {
            code |= bits(1);
            int count = h.count[len];
            if (code - count < first) return h.symbol[index + (code - first)];
            index += count;
            first += count;
            first <<= 1;
            code <<= 1;
        }
        throw std::runtime_error("zip: ongeldige code");
    }

    void stored() {
        bitbuf_ = 0;
        bitcnt_ = 0;
        if (pos_ + 4 > n_) throw std::runtime_error("zip: gegevens afgekapt");
        unsigned len = in_[pos_] | (in_[pos_ + 1] << 8);
        pos_ += 4;
        if (pos_ + len > n_) throw std::runtime_error("zip: gegevens afgekapt");
        out_.insert(out_.end(), in_ + pos_, in_ + pos_ + len);
        pos_ += len;
    }

    void codes(const Huffman& lencode, const Huffman& distcode) {
        static const short lbase[29] = {3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31,
                                        35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258};
        static const short lext[29] = {0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2,
                                       3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0};
        static const short dbase[30] = {1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193,
                                        257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145,
                                        8193, 12289, 16385, 24577};
        static const short dext[30] = {0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6,
                                       7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13};
        for (;;) {
            int sym = decode(lencode);
            if (sym < 256) {
                out_.push_back(uint8_t(sym));
            } else if (sym == 256) {
                return;
            } else {
                sym -= 257;
                if (sym >= 29) throw std::runtime_error("zip: ongeldige lengte");
                int len = lbase[sym] + bits(lext[sym]);
                int ds = decode(distcode);
                if (ds >= 30) throw std::runtime_error("zip: ongeldige afstand");
                size_t dist = size_t(dbase[ds] + bits(dext[ds]));
                if (dist > out_.size()) throw std::runtime_error("zip: afstand te ver");
                size_t from = out_.size() - dist;
                for (int k = 0; k < len; ++k) {
                    uint8_t b = out_[from + k];
                    out_.push_back(b);
                }
            }
        }
    }

    void fixed() {
        static Huffman lencode, distcode;
        static bool built = false;
        if (!built) {
            short lengths[288];
            int s = 0;
            for (; s < 144; ++s) lengths[s] = 8;
            for (; s < 256; ++s) lengths[s] = 9;
            for (; s < 280; ++s) lengths[s] = 7;
            for (; s < 288; ++s) lengths[s] = 8;
            construct(lencode, lengths, 288);
            for (s = 0; s < 30; ++s) lengths[s] = 5;
            construct(distcode, lengths, 30);
            built = true;
        }
        codes(lencode, distcode);
    }

    void dynamic() {
        static const short order[19] = {16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15};
        short lengths[320] = {0};
        int nlen = bits(5) + 257, ndist = bits(5) + 1, ncode = bits(4) + 4;
        if (nlen > 286 || ndist > 30) throw std::runtime_error("zip: ongeldige tabel");
        int idx;
        for (idx = 0; idx < ncode; ++idx) lengths[order[idx]] = short(bits(3));
        for (; idx < 19; ++idx) lengths[order[idx]] = 0;
        Huffman lencode, distcode;
        if (construct(lencode, lengths, 19) != 0) throw std::runtime_error("zip: ongeldige tabel");
        idx = 0;
        while (idx < nlen + ndist) {
            int sym = decode(lencode);
            if (sym < 16) {
                lengths[idx++] = short(sym);
            } else {
                short len = 0;
                if (sym == 16) {
                    if (idx == 0) throw std::runtime_error("zip: ongeldige tabel");
                    len = lengths[idx - 1];
                    sym = 3 + bits(2);
                } else if (sym == 17) {
                    sym = 3 + bits(3);
                } else {
                    sym = 11 + bits(7);
                }
                if (idx + sym > nlen + ndist) throw std::runtime_error("zip: ongeldige tabel");
                while (sym--) lengths[idx++] = len;
            }
        }
        if (lengths[256] == 0) throw std::runtime_error("zip: ongeldige tabel");
        int err = construct(lencode, lengths, nlen);
        if (err < 0 || (err > 0 && nlen - lencode.count[0] != 1)) throw std::runtime_error("zip: ongeldige tabel");
        err = construct(distcode, lengths + nlen, ndist);
        if (err < 0 || (err > 0 && ndist - distcode.count[0] != 1)) throw std::runtime_error("zip: ongeldige tabel");
        codes(lencode, distcode);
    }
};

// Een zip-archief openen en bestanden eruit halen
class ZipReader {
public:
    explicit ZipReader(const fs::path& path) {
        std::ifstream f(path, std::ios::binary);
        if (!f) throw ExitError("Kan " + path.string() + " niet openen.");
        data_.assign(std::istreambuf_iterator<char>(f), std::istreambuf_iterator<char>());
        if (data_.size() < 22) throw std::runtime_error("geen geldig xlsx-bestand: " + path.string());
        size_t eocd = string::npos;
        for (size_t p = data_.size() - 22 + 1; p-- > 0 && data_.size() - p <= 65557;)
            if (u32(p) == 0x06054b50) { eocd = p; break; }
        if (eocd == string::npos) throw std::runtime_error("geen geldig xlsx-bestand: " + path.string());
        size_t count = u16(eocd + 10), p = u32(eocd + 16);
        for (size_t k = 0; k < count; ++k) {
            if (p + 46 > data_.size() || u32(p) != 0x02014b50) throw std::runtime_error("zip: beschadigde inhoudsopgave");
            Entry e;
            e.method = u16(p + 10);
            e.crc = u32(p + 16);
            e.csize = u32(p + 20);
            e.usize = u32(p + 24);
            size_t nlen = u16(p + 28), xlen = u16(p + 30), clen = u16(p + 32);
            e.offset = u32(p + 42);
            string name(reinterpret_cast<const char*>(&data_[p + 46]), nlen);
            entries_[name] = e;
            p += 46 + nlen + xlen + clen;
        }
    }

    bool has(const string& name) const { return entries_.count(name) > 0; }

    string read(const string& name) const {
        auto it = entries_.find(name);
        if (it == entries_.end()) throw std::runtime_error("ontbrekend onderdeel in de xlsx: " + name);
        const Entry& e = it->second;
        size_t p = e.offset;
        if (p + 30 > data_.size() || u32(p) != 0x04034b50) throw std::runtime_error("zip: beschadigd bestand " + name);
        size_t start = p + 30 + u16(p + 26) + u16(p + 28);
        if (start + e.csize > data_.size()) throw std::runtime_error("zip: beschadigd bestand " + name);
        vector<uint8_t> out;
        out.reserve(e.usize);
        if (e.method == 0) {
            out.assign(data_.begin() + start, data_.begin() + start + e.csize);
        } else if (e.method == 8) {
            Inflater(&data_[start], e.csize, out).run();
        } else {
            throw std::runtime_error("zip: onbekende compressie in " + name);
        }
        if (out.size() != e.usize || crc32(out.data(), out.size()) != e.crc)
            throw std::runtime_error("zip: controle mislukt voor " + name);
        return string(out.begin(), out.end());
    }

private:
    struct Entry {
        uint32_t method = 0, crc = 0, csize = 0, usize = 0, offset = 0;
    };
    vector<uint8_t> data_;
    std::map<string, Entry> entries_;
    uint32_t u16(size_t p) const { return data_[p] | (data_[p + 1] << 8); }
    uint32_t u32(size_t p) const {
        return uint32_t(data_[p]) | (uint32_t(data_[p + 1]) << 8) | (uint32_t(data_[p + 2]) << 16) |
               (uint32_t(data_[p + 3]) << 24);
    }
};

// Een eenvoudige XML-lezer: geeft na elkaar begin-tags, eind-tags en tekst terug
void append_utf8(string& out, unsigned long cp) {
    if (cp < 0x80) out += char(cp);
    else if (cp < 0x800) { out += char(0xC0 | (cp >> 6)); out += char(0x80 | (cp & 0x3F)); }
    else if (cp < 0x10000) { out += char(0xE0 | (cp >> 12)); out += char(0x80 | ((cp >> 6) & 0x3F)); out += char(0x80 | (cp & 0x3F)); }
    else { out += char(0xF0 | (cp >> 18)); out += char(0x80 | ((cp >> 12) & 0x3F)); out += char(0x80 | ((cp >> 6) & 0x3F)); out += char(0x80 | (cp & 0x3F)); }
}

string xml_unescape(const string& s) {
    string out;
    out.reserve(s.size());
    for (size_t i = 0; i < s.size(); ++i) {
        char c = s[i];
        if (c == '\r') {                                       // XML: regeleinden worden \n
            out += '\n';
            if (i + 1 < s.size() && s[i + 1] == '\n') ++i;
            continue;
        }
        if (c != '&') { out += c; continue; }
        size_t semi = s.find(';', i);
        if (semi == string::npos) { out += c; continue; }
        string ent = s.substr(i + 1, semi - i - 1);
        if (ent == "amp") out += '&';
        else if (ent == "lt") out += '<';
        else if (ent == "gt") out += '>';
        else if (ent == "quot") out += '"';
        else if (ent == "apos") out += '\'';
        else if (!ent.empty() && ent[0] == '#')
            append_utf8(out, ent.size() > 1 && (ent[1] == 'x' || ent[1] == 'X') ? std::strtoul(ent.c_str() + 2, nullptr, 16)
                                                                              : std::strtoul(ent.c_str() + 1, nullptr, 10));
        else { out += c; continue; }
        i = semi;
    }
    return out;
}

struct XmlItem {
    bool is_text = false, is_end = false, is_empty = false;
    string name, text;                                         // name zonder voorvoegsel (x:c -> c)
    vector<std::pair<string, string>> attrs;
    string attr(const string& key) const {                    // attr("id") vindt ook "r:id"
        for (const auto& a : attrs)
            if (a.first == key) return a.second;
        for (const auto& a : attrs)
            if (a.first.size() > key.size() && a.first.compare(a.first.size() - key.size() - 1, string::npos, ":" + key) == 0)
                return a.second;
        return "";
    }
    bool has(const string& key) const {
        for (const auto& a : attrs)
            if (a.first == key) return true;
        return false;
    }
};

class XmlReader {
public:
    explicit XmlReader(const string& x) : x_(x) {}

    bool next(XmlItem& it) {
        it = XmlItem();
        while (p_ < x_.size()) {
            if (x_[p_] != '<') {
                size_t q = x_.find('<', p_);
                if (q == string::npos) q = x_.size();
                it.is_text = true;
                it.text = xml_unescape(x_.substr(p_, q - p_));
                p_ = q;
                return true;
            }
            if (x_.compare(p_, 4, "<!--") == 0) {
                size_t q = x_.find("-->", p_);
                p_ = q == string::npos ? x_.size() : q + 3;
                continue;
            }
            if (x_.compare(p_, 9, "<![CDATA[") == 0) {
                size_t q = x_.find("]]>", p_);
                if (q == string::npos) q = x_.size();
                it.is_text = true;
                it.text = x_.substr(p_ + 9, q - p_ - 9);
                p_ = std::min(x_.size(), q + 3);
                return true;
            }
            if (p_ + 1 < x_.size() && (x_[p_ + 1] == '?' || x_[p_ + 1] == '!')) {
                size_t q = x_.find('>', p_);
                p_ = q == string::npos ? x_.size() : q + 1;
                continue;
            }
            ++p_;                                              // na '<'
            if (p_ < x_.size() && x_[p_] == '/') {
                it.is_end = true;
                ++p_;
            }
            size_t a = p_;
            while (p_ < x_.size() && !std::isspace((unsigned char)x_[p_]) && x_[p_] != '>' && x_[p_] != '/') ++p_;
            it.name = local(x_.substr(a, p_ - a));
            for (;;) {
                while (p_ < x_.size() && std::isspace((unsigned char)x_[p_])) ++p_;
                if (p_ >= x_.size()) break;
                if (x_[p_] == '>') { ++p_; break; }
                if (x_[p_] == '/') { it.is_empty = true; ++p_; continue; }
                size_t k = p_;
                while (p_ < x_.size() && x_[p_] != '=' && !std::isspace((unsigned char)x_[p_]) && x_[p_] != '>') ++p_;
                string key = x_.substr(k, p_ - k);
                while (p_ < x_.size() && std::isspace((unsigned char)x_[p_])) ++p_;
                string val;
                if (p_ < x_.size() && x_[p_] == '=') {
                    ++p_;
                    while (p_ < x_.size() && std::isspace((unsigned char)x_[p_])) ++p_;
                    char quote = p_ < x_.size() ? x_[p_] : '"';
                    size_t q = x_.find(quote, p_ + 1);
                    if (q == string::npos) q = x_.size();
                    val = xml_unescape(x_.substr(p_ + 1, q - p_ - 1));
                    p_ = std::min(x_.size(), q + 1);
                }
                it.attrs.emplace_back(key, val);
            }
            return true;
        }
        return false;
    }

private:
    const string& x_;
    size_t p_ = 0;
    static string local(const string& n) {
        size_t c = n.find(':');
        return c == string::npos ? n : n.substr(c + 1);
    }
};

// Een tabblad: cellen per (rij, kolom), zoals ws.cell(r, c).value in openpyxl
struct Sheet {
    string title;
    std::unordered_map<long long, Value> cells;
    int max_row = 0;

    const Value& cell(int r, int c) const {
        static const Value none;
        auto it = cells.find(key(r, c));
        return it == cells.end() ? none : it->second;
    }

    int find_row(int col, const string& starts_with_, int first = 1) const {
        for (int r = first; r <= max_row; ++r) {
            const Value& v = cell(r, col);
            if (v.is_str() && starts_with(strip(v.s), starts_with_)) return r;
        }
        throw std::runtime_error("'" + starts_with_ + "' niet gevonden in '" + title + "'");
    }

    static long long key(int r, int c) { return (long long)r * 100000 + c; }
};

// openpyxl: is dit getalformaat een datum/tijd?
bool is_date_format(string fmt) {
    fmt = fmt.substr(0, fmt.find(';'));
    string clean;
    for (size_t i = 0; i < fmt.size();) {
        if (fmt[i] == '"') {                                   // tekst tussen aanhalingstekens negeren
            size_t q = fmt.find('"', i + 1);
            if (q == string::npos) { clean += fmt.substr(i); break; }
            i = q + 1;
        } else if (fmt[i] == '[') {                            // [..] negeren, behalve [h], [mm], [ss], ...
            size_t q = fmt.find(']', i + 1);
            if (q == string::npos) { clean += fmt.substr(i); break; }
            string in = fmt.substr(i + 1, q - i - 1);
            static const std::set<string> keep = {"h", "hh", "m", "mm", "s", "ss"};
            if (keep.count(in)) clean += fmt.substr(i, q - i + 1);
            i = q + 1;
        } else {
            clean += fmt[i++];
        }
    }
    for (size_t i = 0; i < clean.size(); ++i)
        if (std::strchr("dmhysDMHYS", clean[i]) && clean[i] != '\0' && (i == 0 || (clean[i - 1] != '_' && clean[i - 1] != '\\')))
            return true;
    return false;
}

class Workbook {
public:
    explicit Workbook(const fs::path& path) : zip_(path) {
        // 1. tabbladen en waar ze in het archief zitten
        std::map<string, string> rels;
        {
            string x = zip_.read("xl/_rels/workbook.xml.rels");
            XmlReader rd(x);
            XmlItem it;
            while (rd.next(it))
                if (!it.is_text && !it.is_end && it.name == "Relationship") {
                    string target = it.attr("Target");
                    if (!target.empty() && target[0] == '/') target = target.substr(1);
                    else target = "xl/" + target;
                    rels[it.attr("Id")] = target;
                }
        }
        {
            string x = zip_.read("xl/workbook.xml");
            XmlReader rd(x);
            XmlItem it;
            while (rd.next(it))
                if (!it.is_text && !it.is_end && it.name == "sheet") sheets_.emplace_back(it.attr("name"), rels[it.attr("id")]);
        }
        // 2. gedeelde teksten
        if (zip_.has("xl/sharedStrings.xml")) {
            string x = zip_.read("xl/sharedStrings.xml");
            XmlReader rd(x);
            XmlItem it;
            bool in_si = false, in_t = false, in_rph = false;
            string cur;
            while (rd.next(it)) {
                if (it.is_text) {
                    if (in_si && in_t && !in_rph) cur += it.text;
                } else if (it.name == "si") {
                    if (it.is_end) { shared_.push_back(cur); in_si = false; }
                    else if (it.is_empty) shared_.push_back("");
                    else { in_si = true; cur.clear(); }
                } else if (it.name == "rPh") {
                    in_rph = !it.is_end && !it.is_empty;
                } else if (it.name == "t") {
                    in_t = !it.is_end && !it.is_empty;
                }
            }
        }
        // 3. welke celstijlen een datum/tijd-formaat hebben
        if (zip_.has("xl/styles.xml")) {
            string x = zip_.read("xl/styles.xml");
            XmlReader rd(x);
            XmlItem it;
            std::map<int, string> fmts;
            vector<int> xf_fmt;
            bool in_xfs = false;
            while (rd.next(it)) {
                if (it.is_text) continue;
                if (it.name == "numFmt" && !it.is_end) fmts[std::atoi(it.attr("numFmtId").c_str())] = it.attr("formatCode");
                else if (it.name == "cellXfs") in_xfs = !it.is_end && !it.is_empty;
                else if (it.name == "xf" && !it.is_end && in_xfs) xf_fmt.push_back(std::atoi(it.attr("numFmtId").c_str()));
            }
            for (int id : xf_fmt) {
                bool date = (id >= 14 && id <= 22) || (id >= 45 && id <= 47);
                if (id >= 164 && fmts.count(id)) date = is_date_format(fmts[id]);
                xf_is_date_.push_back(date);
            }
        }
    }

    bool has_sheet(const string& name) const {
        for (const auto& s : sheets_)
            if (s.first == name) return true;
        return false;
    }

    const Sheet& sheet(const string& name) {                   // wb[name]
        auto c = cache_.find(name);
        if (c != cache_.end()) return c->second;
        for (const auto& s : sheets_)
            if (s.first == name) return cache_[name] = parse_sheet(name, zip_.read(s.second));
        throw std::runtime_error("tabblad '" + name + "' bestaat niet in de Excel");
    }

private:
    ZipReader zip_;
    vector<std::pair<string, string>> sheets_;
    vector<string> shared_;
    vector<bool> xf_is_date_;
    std::map<string, Sheet> cache_;

    static void ref_to_rc(const string& ref, int& r, int& c) {   // "AB12" -> rij 12, kolom 28
        c = 0;
        size_t i = 0;
        for (; i < ref.size() && std::isalpha((unsigned char)ref[i]); ++i) c = c * 26 + (std::toupper((unsigned char)ref[i]) - 'A' + 1);
        r = std::atoi(ref.c_str() + i);
    }

    Value make_value(const string& t, int style, bool has_v, const string& v, bool has_is, const string& is_text) const {
        Value val;
        if (t == "inlineStr") {
            if (has_is) { val.kind = Value::STR; val.s = is_text; }
            return val;
        }
        if (!has_v || v.empty()) return val;                  // geen waarde -> None
        if (t == "s") {
            size_t k = std::stoul(v);
            if (k >= shared_.size()) throw std::runtime_error("ongeldige tekstverwijzing in de Excel");
            val.kind = Value::STR;
            val.s = shared_[k];
        } else if (t == "str" || t == "e" || t == "d") {
            val.kind = Value::STR;
            val.s = v;
        } else if (t == "b") {
            val.kind = Value::BOOL;
            val.i = std::atoi(v.c_str()) != 0;
        } else {                                               // getal
            bool is_float = v.find_first_of(".eE") != string::npos;
            double d = std::strtod(v.c_str(), nullptr);
            if (style >= 0 && style < (int)xf_is_date_.size() && xf_is_date_[style]) {
                double frac = d - std::floor(d);                // tijdstip: uren uit het deel na de komma
                val.kind = Value::TIME;
                val.i = (long long)std::nearbyint(frac * 86400.0 * 1000.0) % 86400000LL;
            } else if (is_float) {
                val.kind = Value::FLOAT;
                val.f = d;
            } else {
                val.kind = Value::INT;
                val.i = std::stoll(v);
            }
        }
        return val;
    }

    Sheet parse_sheet(const string& name, const string& x) const {
        Sheet sh;
        sh.title = name;
        XmlReader rd(x);
        XmlItem it;
        int row = 0, col = 0, r = 0, c = 0, style = 0;
        string t, v, is_text;
        bool in_c = false, in_v = false, has_v = false, in_is = false, in_t = false, in_rph = false, has_is = false;
        auto store = [&]() {
            sh.cells[Sheet::key(r, c)] = make_value(t, style, has_v, v, has_is, is_text);
            sh.max_row = std::max(sh.max_row, r);
        };
        while (rd.next(it)) {
            if (it.is_text) {
                if (in_v) v += it.text;
                else if (in_is && in_t && !in_rph) is_text += it.text;
                continue;
            }
            if (it.name == "row" && !it.is_end) {
                string rr = it.attr("r");
                row = rr.empty() ? row + 1 : std::atoi(rr.c_str());
                col = 0;
            } else if (it.name == "c" && !it.is_end) {
                string ref = it.attr("r");
                if (ref.empty()) { r = row; c = col + 1; }
                else ref_to_rc(ref, r, c);
                col = c;
                t = it.has("t") ? it.attr("t") : "n";
                style = it.has("s") ? std::atoi(it.attr("s").c_str()) : 0;
                v.clear(); is_text.clear();
                has_v = has_is = false;
                if (it.is_empty) store();
                else in_c = true;
            } else if (it.name == "c" && it.is_end) {
                if (in_c) store();
                in_c = false;
            } else if (in_c && it.name == "v") {
                if (it.is_end) in_v = false;
                else { has_v = true; in_v = !it.is_empty; }
            } else if (in_c && it.name == "is") {
                if (it.is_end) in_is = false;
                else { has_is = true; in_is = !it.is_empty; }
            } else if (in_is && it.name == "rPh") {
                in_rph = !it.is_end && !it.is_empty;
            } else if (in_is && it.name == "t") {
                in_t = !it.is_end && !it.is_empty;
            }
        }
        return sh;
    }
};

// Een xlsx schrijven (voor het overzicht): XML-bestanden in een zip zonder compressie
string xml_escape(const string& s) {
    string out;
    for (unsigned char ch : s) {
        if (ch == '&') out += "&amp;";
        else if (ch == '<') out += "&lt;";
        else if (ch == '>') out += "&gt;";
        else if (ch == '"') out += "&quot;";
        else if (ch < 0x20 && ch != '\t' && ch != '\n' && ch != '\r') continue;
        else out += char(ch);
    }
    return out;
}

class ZipWriter {
public:
    void add(const string& name, const string& content) {
        Entry e{name, crc32(reinterpret_cast<const uint8_t*>(content.data()), content.size()),
                uint32_t(content.size()), uint32_t(buf_.size())};
        u32(0x04034b50); u16(20); u16(0); u16(0); u16(0); u16(0x21);   // versie, vlaggen, geen compressie, 1 jan 1980
        u32(e.crc); u32(e.size); u32(e.size); u16(uint16_t(name.size())); u16(0);
        buf_.insert(buf_.end(), name.begin(), name.end());
        buf_.insert(buf_.end(), content.begin(), content.end());
        entries_.push_back(e);
    }

    void save(const fs::path& path) {
        uint32_t cd = uint32_t(buf_.size());
        for (const Entry& e : entries_) {
            u32(0x02014b50); u16(20); u16(20); u16(0); u16(0); u16(0); u16(0x21);
            u32(e.crc); u32(e.size); u32(e.size); u16(uint16_t(e.name.size())); u16(0); u16(0); u16(0); u16(0); u32(0);
            u32(e.offset);
            buf_.insert(buf_.end(), e.name.begin(), e.name.end());
        }
        uint32_t cd_size = uint32_t(buf_.size()) - cd;
        u32(0x06054b50); u16(0); u16(0); u16(uint16_t(entries_.size())); u16(uint16_t(entries_.size()));
        u32(cd_size); u32(cd); u16(0);
        std::ofstream f(path, std::ios::binary);
        if (!f) throw std::runtime_error("kan " + path.string() + " niet schrijven (is het bestand open in Excel?)");
        f.write(reinterpret_cast<const char*>(buf_.data()), std::streamsize(buf_.size()));
    }

private:
    struct Entry {
        string name;
        uint32_t crc, size, offset;
    };
    vector<uint8_t> buf_;
    vector<Entry> entries_;
    void u16(uint32_t v) { buf_.push_back(uint8_t(v)); buf_.push_back(uint8_t(v >> 8)); }
    void u32(uint32_t v) { u16(v & 0xFFFF); u16(v >> 16); }
};

string col_letter(int c) {                                     // 1 -> A, 28 -> AB
    string s;
    while (c > 0) { s.insert(s.begin(), char('A' + (c - 1) % 26)); c = (c - 1) / 26; }
    return s;
}

struct SheetOut {
    string title;
    vector<string> header;
    vector<vector<PyVal>> rows;
};

void save_xlsx(const fs::path& path, const vector<SheetOut>& sheets) {
    const string ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const string rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const string head = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
    ZipWriter zip;

    string types = head + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
        "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
        "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
        "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
        "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>";
    for (size_t k = 1; k <= sheets.size(); ++k)
        types += "<Override PartName=\"/xl/worksheets/sheet" + std::to_string(k) +
                 ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
    types += "</Types>";
    zip.add("[Content_Types].xml", types);

    zip.add("_rels/.rels", head + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" Type=\"" + rns + "/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>");

    string wb = head + "<workbook xmlns=\"" + ns + "\" xmlns:r=\"" + rns + "\"><bookViews><workbookView activeTab=\"0\"/></bookViews><sheets>";
    string rels = head + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">";
    for (size_t k = 1; k <= sheets.size(); ++k) {
        string id = std::to_string(k);
        wb += "<sheet name=\"" + xml_escape(sheets[k - 1].title) + "\" sheetId=\"" + id + "\" r:id=\"rId" + id + "\"/>";
        rels += "<Relationship Id=\"rId" + id + "\" Type=\"" + rns + "/worksheet\" Target=\"worksheets/sheet" + id + ".xml\"/>";
    }
    wb += "</sheets></workbook>";
    rels += "<Relationship Id=\"rId" + std::to_string(sheets.size() + 1) + "\" Type=\"" + rns +
            "/styles\" Target=\"styles.xml\"/></Relationships>";
    zip.add("xl/workbook.xml", wb);
    zip.add("xl/_rels/workbook.xml.rels", rels);

    // stijl 0 = gewoon, stijl 1 = kop: vet met grijze achtergrond (Font(bold=True), PatternFill("solid", "EEEEEE"))
    zip.add("xl/styles.xml", head + "<styleSheet xmlns=\"" + ns + "\">"
        "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font>"
        "<font><b/><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>"
        "<fills count=\"3\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill>"
        "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"00EEEEEE\"/></patternFill></fill></fills>"
        "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
        "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
        "<cellXfs count=\"2\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
        "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFont=\"1\" applyFill=\"1\"/></cellXfs>"
        "<cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles></styleSheet>");

    for (size_t k = 0; k < sheets.size(); ++k) {
        const SheetOut& so = sheets[k];
        size_t ncol = so.header.size();
        for (const auto& r : so.rows) ncol = std::max(ncol, r.size());
        size_t nrow = so.rows.size() + 1;

        // kolombreedte: min(langste tekst + 2, 48)
        vector<size_t> width(ncol, 0);
        for (size_t c = 0; c < so.header.size(); ++c) width[c] = std::max(width[c], u8len(so.header[c]));
        for (const auto& r : so.rows)
            for (size_t c = 0; c < r.size(); ++c) width[c] = std::max(width[c], r[c].truthy() ? u8len(r[c].str()) : 0);

        string x = head + "<worksheet xmlns=\"" + ns + "\" xmlns:r=\"" + rns + "\">";
        x += "<dimension ref=\"A1:" + col_letter(int(ncol)) + std::to_string(nrow) + "\"/>";
        x += "<sheetViews><sheetView workbookViewId=\"0\">"            // bovenste rij vast (freeze_panes = "A2")
             "<pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/>"
             "<selection pane=\"bottomLeft\" activeCell=\"A1\" sqref=\"A1\"/></sheetView></sheetViews>";
        x += "<sheetFormatPr defaultRowHeight=\"15\"/><cols>";
        for (size_t c = 0; c < ncol; ++c)
            x += "<col min=\"" + std::to_string(c + 1) + "\" max=\"" + std::to_string(c + 1) + "\" width=\"" +
                 std::to_string(std::min<size_t>(width[c] + 2, 48)) + "\" customWidth=\"1\"/>";
        x += "</cols><sheetData>";
        auto text_cell = [&](const string& ref, const string& s, bool bold) {
            string st = bold ? " s=\"1\"" : "";
            if (s.empty()) return "<c r=\"" + ref + "\"" + st + "/>";
            return "<c r=\"" + ref + "\"" + st + " t=\"inlineStr\"><is><t xml:space=\"preserve\">" + xml_escape(s) + "</t></is></c>";
        };
        x += "<row r=\"1\">";
        for (size_t c = 0; c < so.header.size(); ++c) x += text_cell(col_letter(int(c + 1)) + "1", so.header[c], true);
        x += "</row>";
        for (size_t r = 0; r < so.rows.size(); ++r) {
            string rn = std::to_string(r + 2);
            x += "<row r=\"" + rn + "\">";
            for (size_t c = 0; c < so.rows[r].size(); ++c) {
                const PyVal& v = so.rows[r][c];
                string ref = col_letter(int(c + 1)) + rn;
                if (v.kind == PyVal::STR) {
                    x += text_cell(ref, v.s, false);
                } else {                                       // getal: zoals openpyxl ("%.16g": 1.0 -> 1)
                    double d = v.num();
                    char num[40] = "";
                    if (std::isfinite(d)) std::snprintf(num, sizeof num, "%.16g", d);
                    x += "<c r=\"" + ref + "\" t=\"n\"><v>" + num + "</v></c>";
                }
            }
            x += "</row>";
        }
        x += "</sheetData><pageMargins left=\"0.75\" right=\"0.75\" top=\"1\" bottom=\"1\" header=\"0.5\" footer=\"0.5\"/></worksheet>";
        zip.add("xl/worksheets/sheet" + std::to_string(k + 1) + ".xml", x);
    }
    zip.save(path);
}


// ============================================================================
// DE GEGEVENS
// ============================================================================

struct Pattern {                 // een werkpatroon: 7 dagen, 1 = werken
    string label;
    vector<int> days;
};

struct Shift {
    string name;
    int start = 0;               // startuur
    long long min = 0;           // minimale bezetting (Case C)
    long long min_buffer = 0;    // eis na buffer (Case D)
};

struct Comp {                    // component = de lijnen met dezelfde (shift, patroon, type)
    string shift;
    int pattern = 0;
    long long type = 0;
    long long lines = 0;
    vector<int> members;         // indexen in de lijst met personen
};

struct Req {
    long long req = 0, min_type1 = 0;
};

struct Pref {                    // de cyclische voorkeur, ontleed ("" = geen)
    std::set<string> no_shift;
    string only_shift;
    bool no_weekend = false;
    string pref_shift;
    std::set<int> days_off, days_on;
};

struct Person {
    string id;
    double rate = 0;             // tewerkstelling (1.0, 0.75, 0.5)
    long long type = 0;          // 1 of 2
    string text;                 // cyclische voorkeur (tekst)
    Pref pref;
    bool is_new = false;         // aanwerving?
    int comp = -1;               // index van de component (lijn) waarop de persoon staat
    long long score = 0;         // strafscore
    string status;
};

using Summary = vector<std::pair<string, PyVal>>;


// ============================================================================
// STAP 0 - INLEZEN
// ============================================================================

vector<Pattern> read_patterns(const Sheet& ws) {
    // De 7 werkpatronen (ma-vr, di-za, ...) uit 'Work patterns used by the model'.
    vector<Pattern> pats;
    string prefix = POLICY.substr(0, 3) + " pattern";
    for (int r = 1; r <= ws.max_row; ++r) {
        const Value& label = ws.cell(r, 2);
        if (label.is_str() && starts_with(label.s, prefix)) {
            Pattern p;
            p.label = label.s;
            for (int c = 3; c < 10; ++c) p.days.push_back(int(py_int(ws.cell(r, c), "werkpatroon rij " + std::to_string(r))));
            pats.push_back(p);
        }
    }
    if (pats.size() != 7) throw std::runtime_error("verwachtte 7 werkpatronen");
    return pats;
}

vector<Shift> read_shifts(const Sheet& ws, char dept) {
    // Shiften met startuur en minimale bezetting (Case C) en eis na buffer (Case D).
    int d = int(string("ABCD").find(dept));
    int r0 = ws.find_row(2, "Staffing requirements per shift from Case C");
    vector<Shift> shifts;
    for (int r = r0 + 2; r < r0 + 12; ++r) {
        const Value& name = ws.cell(r, 5);
        const Value& start = ws.cell(r, 4);
        if (ws.cell(r, 2) != SCENARIO || name.kind == Value::NONE || name == "-" || start.kind == Value::NONE || start == "-")
            continue;
        int hour;
        if (start.kind == Value::TIME) hour = start.hour();
        else hour = int(py_int_from_string(start.str().substr(0, start.str().find(':')), "startuur rij " + std::to_string(r)));
        Shift s;
        s.name = name.str();
        s.start = hour;
        s.min = py_int(ws.cell(r, 6 + d), "minimum rij " + std::to_string(r));
        s.min_buffer = py_int(ws.cell(r, 10 + d), "eis met buffer rij " + std::to_string(r));
        shifts.push_back(s);
    }
    std::stable_sort(shifts.begin(), shifts.end(), [](const Shift& a, const Shift& b) { return a.start < b.start; });
    return shifts;
}

void read_lines(const Sheet& ws, char dept, vector<Comp>& comps, std::map<string, Req>& req) {
    // De lijnen van Case D per (shift, patroon, type). Elk zo'n groepje heet een component.
    int r1 = ws.find_row(2, "Model " + SCENARIO);
    int r2 = ws.find_row(2, "Total number of nurses (" + SCENARIO + ")", r1);
    for (int r = r1; r < r2; ++r) {
        if (ws.cell(r, 2) != POLICY || ws.cell(r, 3) != string("Dept ") + dept) continue;
        string shift = ws.cell(r, 5).str();
        string where = "rij " + std::to_string(r);
        req[shift] = Req{py_round(ws.cell(r, 22), "Req m " + where), py_round(ws.cell(r, 23), "min type I " + where)};
        for (int k = 0; k < 7; ++k) {
            const int typ[2] = {1, 2}, col[2] = {6 + k, 13 + k};   // F..L = type I, M..S = type II
            for (int j = 0; j < 2; ++j) {
                const Value& v = ws.cell(r, col[j]);
                long long lines = v.truthy() ? py_round(v, "lijnen " + where) : 0;
                if (lines) {
                    Comp c;
                    c.shift = shift;
                    c.pattern = k;
                    c.type = typ[j];
                    c.lines = lines;
                    comps.push_back(c);
                }
            }
        }
    }
}

Pref parse_preference(const string& text);

vector<Person> read_nurses(Workbook& wb, char dept) {
    const Sheet& ws = wb.sheet(string("Department ") + dept + " - Nurse Char.");
    vector<Person> nurses;
    for (int r = 1; r <= ws.max_row; ++r) {
        const Value& nid = ws.cell(r, 1);
        if (nid.is_str() && std::regex_match(strip(nid.s), ID_RE)) {
            const Value& t = ws.cell(r, 18);
            Person n;
            n.id = strip(nid.s);
            n.text = strip(t.truthy() ? t.str() : "");
            n.rate = py_float(ws.cell(r, 16), "tewerkstelling " + n.id);
            n.type = py_int(ws.cell(r, 17), "type " + n.id);
            n.pref = parse_preference(n.text);
            n.is_new = false;
            nurses.push_back(n);
        }
    }
    return nurses;
}

std::map<string, vector<long long>> read_monthly_preferences(Workbook& wb, char dept) {
    const Sheet& ws = wb.sheet(string("Department ") + dept + " - Preferences");
    int head = ws.find_row(2, "E");
    const char* expect[5] = {"E", "D", "L", "N", "F"};
    for (int c = 2; c < 7; ++c)
        if (ws.cell(head, c) != expect[c - 2]) throw std::runtime_error("onverwachte kop in '" + ws.title + "'");
    std::map<string, vector<long long>> prefs;
    for (int r = head + 1; r <= ws.max_row; ++r) {
        const Value& nid = ws.cell(r, 1);
        if (nid.is_str() && std::regex_match(strip(nid.s), ID_RE)) {
            vector<long long> vals;
            for (int c = 2; c < 2 + 140; ++c) {
                const Value& v = ws.cell(r, c);
                vals.push_back(v.is_number() ? py_int(v, "voorkeur") : 0);
            }
            prefs[strip(nid.s)] = vals;
        }
    }
    return prefs;
}


// ============================================================================
// STAP 1 - STRAFSCORE: hoe slecht past een verpleegkundige bij een component?
// ============================================================================

string first_match(const string& t, const std::regex& re) {   // re.search(...).group(1), "" als er niets is
    std::smatch m;
    return std::regex_search(t, m, re) ? capitalize(m[1].str()) : "";
}

Pref parse_preference(const string& text) {
    string t = lower(text);
    Pref p;
    p.no_weekend = contains(t, "no weekend");
    for (const char* s : {"early", "day", "late", "night"})
        if (contains(t, string("no ") + s + " shift")) p.no_shift.insert(capitalize(s));
    static const std::regex only_re("only (early|day|late|night) shifts");
    static const std::regex pref_re("preference for (early|day|late|night) shifts");
    p.only_shift = first_match(t, only_re);
    p.pref_shift = first_match(t, pref_re);
    size_t a = 0;
    for (;;) {                                                // for clause in t.split(","):
        size_t b = t.find(',', a);
        string clause = t.substr(a, b == string::npos ? string::npos : b - a);
        std::set<int> days;
        for (int i = 0; i < 7; ++i)
            if (contains(clause, DAY_NAMES[i])) days.insert(i);
        if (contains(clause, "weekend") && !contains(clause, "no weekend")) days.insert(WEEKEND.begin(), WEEKEND.end());
        if (contains(clause, "day off")) p.days_off.insert(days.begin(), days.end());
        else if (contains(clause, "day on")) p.days_on.insert(days.begin(), days.end());
        if (b == string::npos) break;
        a = b + 1;
    }
    return p;
}

std::optional<long long> score(const Person& nurse, const Comp& comp, const vector<Pattern>& patterns) {
    // leeg (None) = mag niet (type of contract). Anders: aantal strafpunten (0 = past perfect).
    if (nurse.type == 2 && comp.type == 1) return std::nullopt;
    const Pref& p = nurse.pref;
    const vector<int>& days = patterns[comp.pattern].days;
    if (p.no_shift.count(comp.shift) || (!p.only_shift.empty() && comp.shift != p.only_shift)) return std::nullopt;
    if (p.no_weekend) {
        for (int d : WEEKEND)
            if (days[d]) return std::nullopt;
    }
    long long s = (!p.pref_shift.empty() && comp.shift != p.pref_shift) ? PEN_SHIFT_PREF : 0;
    for (int d : p.days_off) s += PEN_DAY_OFF * days[d];
    for (int d : p.days_on) s += PEN_DAY_ON * (1 - days[d]);
    if (nurse.type == 1 && comp.type == 2) s += PEN_TYPE1_ON_TYPE2;
    return s;
}


// ============================================================================
// STAP 2 - DE EIGEN VERPLEEGKUNDIGEN KRIJGEN ELK EEN LIJN
// ============================================================================

void assign_own(vector<Person>& people, vector<Comp>& comps, const vector<Pattern>& patterns) {
    // Herhaal tot iedereen een lijn heeft: kies het koppel (verpleegkundige, vrije lijn)
    // met de laagste strafscore. Bij gelijke score gaat wie het minst perfect passende
    // lijnen heeft eerst, zodat flexibele mensen de rest opvullen. Type I-lijnen eerst.
    size_t n = people.size(), m = comps.size();
    {
        std::set<string> seen;
        for (const Person& p : people)
            if (!seen.insert(p.id).second) throw std::runtime_error("verpleegkundige " + p.id + " staat dubbel in de Excel");
    }
    vector<vector<std::optional<long long>>> S(n, vector<std::optional<long long>>(m));
    for (size_t a = 0; a < n; ++a)
        for (size_t i = 0; i < m; ++i) S[a][i] = score(people[a], comps[i], patterns);
    vector<int> free_;                                         // nog zonder lijn, in volgorde van de Excel
    for (size_t a = 0; a < n; ++a) free_.push_back(int(a));

    auto place = [&](int a, int i, const string& status) {
        free_.erase(std::find(free_.begin(), free_.end(), a));
        comps[i].members.push_back(a);
        people[a].comp = i;
        people[a].score = *S[a][i];
        people[a].status = status;
    };

    for (int phase = 1; phase <= 2; ++phase) {
        while (!free_.empty()) {
            vector<int> open_i;
            for (size_t i = 0; i < m; ++i)
                if (comps[i].type == phase && (long long)comps[i].members.size() < comps[i].lines) open_i.push_back(int(i));
            std::map<int, int> flex;
            for (int a : free_) {
                int f = 0;
                for (int i : open_i)
                    if (S[a][i] && *S[a][i] == 0) ++f;
                flex[a] = f;
            }
            bool found = false;
            std::tuple<long long, int, string, int> best;
            int best_a = -1;
            for (int i : open_i)
                for (int a : free_)
                    if (S[a][i]) {
                        auto option = std::make_tuple(*S[a][i], flex[a], people[a].id, i);
                        if (!found || option < best) { best = option; best_a = a; found = true; }
                    }
            if (!found) break;
            place(best_a, std::get<3>(best), "eigen personeel");
        }
    }

    // wie op geen enkele vrije lijn past (bv. 'No weekend shifts' en alle ma-vr-lijnen bezet),
    // komt als extra persoon op haar best passende lijn
    vector<int> rest = free_;
    std::sort(rest.begin(), rest.end(), [&](int x, int y) { return people[x].id < people[y].id; });
    for (int a : rest) {
        bool found = false;
        std::pair<long long, int> best;
        for (size_t i = 0; i < m; ++i)
            if (S[a][i]) {
                auto option = std::make_pair(*S[a][i], int(i));
                if (!found || option < best) { best = option; found = true; }
            }
        if (!found) throw std::runtime_error(people[a].id + " past op geen enkele lijn");
        place(a, best.second, "eigen personeel (extra op lijn)");
    }
}


// ============================================================================
// STAP 3 - AANWERVEN VOOR HET TEKORT IN FTE
// ============================================================================

vector<double> contracts(double fte) {
    // 24,5 FTE -> 24 voltijdse + 1 halftijdse.
    vector<double> out;
    double rest = fte;
    while (rest > EPS) {
        double rate;
        if (rest >= 1 - EPS) {
            rate = 1.0;
        } else {
            bool found = false;
            rate = 0;
            for (double r : HIRE_CONTRACTS)
                if (r >= rest - EPS && (!found || r < rate)) { rate = r; found = true; }
            if (!found) throw std::runtime_error("geen passend contract");
        }
        out.push_back(rate);
        rest -= rate;
    }
    return out;
}

double fte(const Comp& c, const vector<Person>& people) {
    double t = 0;
    for (int a : c.members) t += people[a].rate;
    return t;
}

vector<int> hire(vector<Person>& people, vector<Comp>& comps, char dept) {
    // Tekort = lijnen - som van de tewerkstelling van het eigen personeel.
    // Eerst type I (als er te weinig type I-FTE is), dan type II.
    // Elke aanwerving gaat naar de component met het grootste FTE-tekort.
    size_t n_nurses = people.size();
    long long need = 0, need_1 = 0;
    for (const Comp& c : comps) {
        need += c.lines;
        if (c.type == 1) need_1 += c.lines;
    }
    double have = 0, have_1 = 0;
    for (size_t a = 0; a < n_nurses; ++a) {
        have += people[a].rate;
        if (people[a].type == 1) have_1 += people[a].rate;
    }
    double hire_1 = std::max(0.0, double(need_1) - have_1);
    double hire_2 = std::max(0.0, double(need) - have - hire_1);

    vector<int> hires;
    const std::pair<int, double> rounds[2] = {{1, hire_1}, {2, hire_2}};
    for (const auto& rnd : rounds) {
        int typ = rnd.first;
        for (double rate : contracts(rnd.second)) {
            char id[32];
            std::snprintf(id, sizeof id, "NEW%c%03d", dept, int(hires.size() + 1));
            Person p;
            p.id = id;
            p.rate = rate;
            p.type = typ;
            p.text = "nieuwe aanwerving";
            p.pref = parse_preference("");
            p.is_new = true;
            p.score = 0;
            p.status = "aanwerving";
            // grootste FTE-tekort eerst; bij gelijk tekort: eigen type, dan een lijn die nog leeg is
            int best = -1;
            std::tuple<double, bool, long long, int> best_key;
            for (size_t i = 0; i < comps.size(); ++i) {
                if (!(typ == 1 || comps[i].type == 2)) continue;
                auto key = std::make_tuple(double(comps[i].lines) - fte(comps[i], people), comps[i].type == typ,
                                           comps[i].lines - (long long)comps[i].members.size(), -int(i));
                if (best < 0 || key > best_key) { best = int(i); best_key = key; }
            }
            if (best < 0) throw std::runtime_error("geen lijn voor een aanwerving van type " + std::to_string(typ));
            people.push_back(p);
            int a = int(people.size() - 1);
            comps[best].members.push_back(a);
            people[a].comp = best;
            hires.push_back(a);
        }
    }
    return hires;
}


// ============================================================================
// STAP 4 - UITVOER
// ============================================================================

std::ofstream open_text(const fs::path& path) {               // tekstbestand (zoals open(path, "w") in Python)
    std::ofstream f(path);
    if (!f) throw std::runtime_error("kan " + path.string() + " niet schrijven");
    return f;
}

std::map<string, int> shift_codes(const vector<Shift>& shifts) {   // 0 = vrij, 1..n = shift op startuur
    std::map<string, int> code;
    for (size_t k = 0; k < shifts.size(); ++k) code[shifts[k].name] = int(k + 1);
    return code;
}

int code_of(const std::map<string, int>& code, const string& shift) {
    auto it = code.find(shift);
    if (it == code.end()) throw std::runtime_error("shift '" + shift + "' ontbreekt in de shifttabel");
    return it->second;
}

vector<int> week_of(const Person& n, const vector<Comp>& comps, const vector<Pattern>& patterns,
                    const std::map<string, int>& code) {          // 28 dagen: de standaardweek x WEEKS
    const Comp& c = comps[n.comp];
    vector<int> d28;
    for (int w = 0; w < WEEKS; ++w)
        for (int d = 0; d < 7; ++d) d28.push_back(patterns[c.pattern].days[d] ? code_of(code, c.shift) : 0);
    return d28;
}

void write_template_files(const fs::path& out, char dept, const vector<Person>& roster, const vector<Comp>& comps,
                          const vector<Pattern>& patterns, const vector<Shift>& shifts,
                          const std::map<string, vector<long long>>& monthly) {
    fs::create_directories(out);
    std::smatch m;
    std::regex_search(SCENARIO, m, std::regex("(\\d+)h"));
    int length = std::stoi(m[1].str());
    auto code = shift_codes(shifts);
    string D(1, dept);

    {
        auto f = open_text(out / ("Shift_system_dpt_" + D + ".txt"));
        f << shifts.size() << "\t" << length << "\n";
        for (const Shift& s : shifts) f << s.start << "\n";
        vector<string> mins;
        for (const Shift& s : shifts) mins.push_back(std::to_string(REQ_WITH_BUFFER ? s.min_buffer : s.min));
        f << join(mins, "\t") << "\n";
    }
    {
        auto f = open_text(out / ("Cyclic_roster_dpt_" + D + ".txt"));
        f << roster.size() << "\n";
        for (const Person& n : roster) {
            vector<string> cells;
            for (int v : week_of(n, comps, patterns, code)) cells.push_back(std::to_string(v));
            f << join(cells, "\t") << "\n";
        }
    }
    {
        auto f = open_text(out / ("Personnel_dpt_" + D + ".txt"));
        for (const Person& n : roster) {
            vector<string> prefs;
            if (n.is_new) {
                prefs.assign(140, std::to_string(HIRE_PREF));
            } else {
                auto it = monthly.find(n.id);
                if (it == monthly.end()) throw std::runtime_error(n.id + " ontbreekt in 'Department " + D + " - Preferences'");
                for (long long v : it->second) prefs.push_back(std::to_string(v));
            }
            f << n.id << "\t" << join(prefs, "\t") << "\t" << py_repr(n.rate) << "\t" << n.type << "\n";
        }
    }
    {
        const ConstraintSettings& c = CONSTRAINTS;
        auto f = open_text(out / ("Constraints_dpt_" + D + ".txt"));
        f << c.min_ass << " " << c.max_ass << "\n" << c.min_cons_wrk << " " << c.max_cons_wrk << "\n";
        for (size_t k = 0; k < shifts.size(); ++k) f << c.min_cons_shift << " " << c.max_cons_shift << "\n";
        for (size_t k = 0; k < shifts.size(); ++k) f << c.min_shift << " " << c.max_shift << "\n";
        f << c.identical << "\n";
    }
}

void write_overview(const fs::path& path, const Summary& summary, const vector<Person>& roster,
                    const vector<Comp>& comps, const vector<Pattern>& patterns, const vector<Shift>& shifts,
                    const std::map<string, Req>& req) {
    vector<SheetOut> sheets;

    SheetOut s1{"Samenvatting", {"", "Waarde"}, {}};
    for (const auto& kv : summary) s1.rows.push_back({PyVal::S(kv.first), kv.second});
    sheets.push_back(s1);

    SheetOut s2{"Cyclisch rooster", {"Verpleegkundige", "Type", "Tewerkstelling", "Cyclische voorkeur", "Shift",
                                     "Patroon", "Strafscore", "Status"}, {}};
    s2.header.insert(s2.header.end(), DAYS_NL.begin(), DAYS_NL.end());
    for (const Person& n : roster) {
        const Comp& c = comps[n.comp];
        const vector<int>& days = patterns[c.pattern].days;
        vector<PyVal> row = {PyVal::S(n.id), PyVal::I(n.type), PyVal::F(n.rate), PyVal::S(n.text), PyVal::S(c.shift),
                             PyVal::S(patterns[c.pattern].label), PyVal::I(n.score), PyVal::S(n.status)};
        for (int d = 0; d < 7; ++d) {
            if (!days[d]) { row.push_back(PyVal::S("-")); continue; }
            auto l = LETTER.find(c.shift);
            if (l == LETTER.end()) throw std::runtime_error("onbekende shift '" + c.shift + "'");
            row.push_back(PyVal::S(l->second));
        }
        s2.rows.push_back(row);
    }
    sheets.push_back(s2);

    SheetOut s3{"Lijnen", {"Shift", "Patroon", "Type", "Lijnen (FTE nodig)", "Personen", "FTE ingevuld",
                           "Verschil (FTE)", "Wie"}, {}};
    for (const Comp& c : comps) {
        vector<double> rates;
        vector<string> ids;
        for (int a : c.members) { rates.push_back(roster[a].rate); ids.push_back(roster[a].id); }
        PyVal f = py_sum(rates);
        PyVal diff = f.kind == PyVal::INT ? PyVal::I(f.i - c.lines) : py_round2(PyVal::F(f.f - double(c.lines)));
        s3.rows.push_back({PyVal::S(c.shift), PyVal::S(patterns[c.pattern].label), PyVal::S(c.type == 1 ? "I" : "II"),
                           PyVal::I(c.lines), PyVal::I((long long)c.members.size()), f, diff, PyVal::S(join(ids, ", "))});
    }
    sheets.push_back(s3);

    SheetOut s4{"Dekking", {"Shift", "Dag", "Personen (cyclisch rooster)", "FTE (na vrije dagen deeltijdsen)",
                            "Type I", "Eis Case D (met buffer)", "Min. type I", "Minimum Case C"}, {}};
    for (const Shift& s : shifts) {
        auto r = req.find(s.name);
        if (r == req.end()) throw std::runtime_error("geen lijnen voor shift '" + s.name + "' in Case D");
        for (int d = 0; d < 7; ++d) {
            vector<double> rates;
            long long n_type1 = 0;
            for (const Person& n : roster)
                if (comps[n.comp].shift == s.name && patterns[comps[n.comp].pattern].days[d]) {
                    rates.push_back(n.rate);
                    if (n.type == 1) ++n_type1;
                }
            s4.rows.push_back({PyVal::S(s.name), PyVal::S(DAYS_NL[d]), PyVal::I((long long)rates.size()), py_sum(rates),
                               PyVal::I(n_type1), PyVal::I(r->second.req), PyVal::I(r->second.min_type1), PyVal::I(s.min)});
        }
    }
    sheets.push_back(s4);
    save_xlsx(path, sheets);
}

void write_readable_roster(const fs::path& path, char dept, const vector<Person>& roster, const vector<Comp>& comps,
                           const vector<Pattern>& patterns, const vector<Shift>& shifts) {
    // Cyclisch rooster met koppen, om zelf te lezen. Zelfde codes als Cyclic_roster_dpt_X.txt.
    auto code = shift_codes(shifts);
    vector<string> days;
    for (int d = 0; d < 7 * WEEKS; ++d) days.push_back(DAYS_NL[d % 7] + std::to_string(d + 1));
    vector<vector<int>> rows;
    for (const Person& n : roster) rows.push_back(week_of(n, comps, patterns, code));

    auto line = [](const string& head, const vector<string>& values) {   // 4 kolommen per dag, '|' tussen de weken
        vector<string> weeks;
        for (int w = 0; w < WEEKS; ++w) {
            vector<string> cells;
            for (int d = w * 7; d < (w + 1) * 7; ++d) cells.push_back(pad_left(values[d], 4));
            weeks.push_back(join(cells, " "));
        }
        return head + " | " + join(weeks, " | ") + "\n";
    };
    auto as_text = [](const vector<int>& v) {
        vector<string> out;
        for (int x : v) out.push_back(std::to_string(x));
        return out;
    };

    vector<string> legend;
    for (const Shift& s : shifts) {
        char b[16];
        std::snprintf(b, sizeof b, "%02d", s.start);
        legend.push_back(std::to_string(code_of(code, s.name)) + " = " + s.name + " (" + b + ":00)");
    }
    size_t own = 0;
    for (const Person& n : roster)
        if (!n.is_new) ++own;
    string D(1, dept);

    auto f = open_text(path);
    f << "Cyclisch rooster afdeling " << D << " (" << SCENARIO << ", " << POLICY << ")\n";
    f << "Codes: 0 = vrij, " << join(legend, ", ") << "\n";
    f << roster.size() << " personen: " << own << " eigen personeel + " << roster.size() - own << " aanwervingen (NEW...)\n";
    f << "Elke week is dezelfde standaardweek (cyclisch rooster = 1 week x 4).\n\n";
    string head = pad_left("Regel", 5) + "  " + pad_right("Persoon", 9) + " " + pad_left("Type", 4) + " " + pad_left("FTE", 5);
    f << line(head, days);
    string dashes(u8len(head) + 3 + WEEKS * (7 * 5 + 2), '-');
    f << dashes << "\n";
    for (size_t k = 0; k < roster.size(); ++k) {                // regelnummer in Cyclic_roster_dpt_X.txt
        const Person& n = roster[k];
        char rate[32];
        std::snprintf(rate, sizeof rate, "%5.2f", n.rate);
        f << line(pad_left(std::to_string(k + 2), 5) + "  " + pad_right(n.id, 9) + " " + pad_left(std::to_string(n.type), 4) +
                      " " + rate, as_text(rows[k]));
    }
    f << dashes << "\n";
    for (const Shift& s : shifts) {
        int k = code_of(code, s.name);
        vector<int> count;
        for (int d = 0; d < 7 * WEEKS; ++d) {
            int c = 0;
            for (const auto& r : rows)
                if (r[d] == k) ++c;
            count.push_back(c);
        }
        f << line(pad_left("", 5) + "  " + pad_right("# " + s.name, 9) + " " + pad_left("min", 4) + " " +
                      pad_left(std::to_string(s.min), 5), as_text(count));
    }
    f << "\n'Regel' = regelnummer van die persoon in files/Cyclic_roster_dpt_" << D << ".txt.\n";
    f << "Onderaan: hoeveel mensen er die dag op die shift staan, tegenover het minimum uit Case C.\n";
}


// ============================================================================
// HOOFDPROGRAMMA
// ============================================================================

Summary run_department(Workbook& wb, char dept, const fs::path& out_dir, const fs::path& overview_dir) {
    if (string("ABCD").find(dept) == string::npos) throw std::runtime_error(string("onbekende afdeling '") + dept + "'");
    const Sheet& ws = wb.sheet(SHEET_CASE_D);
    vector<Pattern> patterns = read_patterns(ws);
    vector<Shift> shifts = read_shifts(ws, dept);
    vector<Comp> comps;
    std::map<string, Req> req;
    read_lines(ws, dept, comps, req);
    vector<Person> people = read_nurses(wb, dept);             // eerst het eigen personeel, daarna de aanwervingen
    auto monthly = read_monthly_preferences(wb, dept);
    size_t n_nurses = people.size();

    assign_own(people, comps, patterns);                       // stap 2
    vector<int> hires = hire(people, comps, dept);              // stap 3
    const vector<Person>& roster = people;

    // controles
    long long need = 0, need_1 = 0;
    for (const Comp& c : comps) {
        need += c.lines;
        if (c.type == 1) need_1 += c.lines;
    }
    vector<double> all_rates, own_rates, hire_rates;
    double total_1 = 0;
    for (size_t a = 0; a < roster.size(); ++a) {
        all_rates.push_back(roster[a].rate);
        (a < n_nurses ? own_rates : hire_rates).push_back(roster[a].rate);
        if (roster[a].type == 1) total_1 += roster[a].rate;
    }
    PyVal total = py_sum(all_rates);
    if (!(total.num() >= double(need) - EPS)) throw std::runtime_error("te weinig FTE");
    if (!(total_1 >= double(need_1) - EPS)) throw std::runtime_error("te weinig type I");
    for (size_t a = 0; a < n_nurses; ++a)
        if (!score(roster[a], comps[roster[a].comp], patterns)) throw std::runtime_error("contractregel geschonden");

    PyVal have = py_sum(own_rates);
    PyVal shortage = have.kind == PyVal::INT ? PyVal::I(need - have.i) : PyVal::F(double(need) - have.f);
    vector<double> distinct = hire_rates;                      // sorted({h["rate"] ...}, reverse=True)
    std::sort(distinct.begin(), distinct.end(), std::greater<double>());
    distinct.erase(std::unique(distinct.begin(), distinct.end()), distinct.end());
    vector<string> parts;
    for (double r : distinct)
        parts.push_back(std::to_string(std::count(hire_rates.begin(), hire_rates.end(), r)) + " x " + py_repr(r));
    long long zero = 0, total_score = 0, extra = 0;
    for (size_t a = 0; a < n_nurses; ++a) {
        if (roster[a].score == 0) ++zero;
        total_score += roster[a].score;
        if (contains(roster[a].status, "extra")) ++extra;
    }

    Summary summary = {
        {"Afdeling", PyVal::S(string(1, dept))},
        {"Lijnen uit Case D (= FTE nodig)", PyVal::I(need)},
        {"   waarvan type I", PyVal::I(need_1)},
        {"Eigen personeel (personen)", PyVal::I((long long)n_nurses)},
        {"Eigen personeel (FTE)", have},
        {"Tekort (FTE; negatief = overschot)", shortage},
        {"Aanwervingen (personen)", PyVal::I((long long)hires.size())},
        {"Aanwervingen (FTE)", py_sum(hire_rates)},
        {"   contracten", PyVal::S(parts.empty() ? "-" : join(parts, ", "))},
        {"Personen in het cyclisch rooster", PyVal::I((long long)roster.size())},
        {"FTE in het cyclisch rooster", total},
        {"Eigen personeel met strafscore 0", PyVal::S(std::to_string(zero) + " van " + std::to_string(n_nurses))},
        {"Totale strafscore eigen personeel", PyVal::I(total_score)},
        {"Eigen personeel extra op een lijn (paste nergens anders)", PyVal::I(extra)},
    };
    string D(1, dept);
    write_template_files(out_dir, dept, roster, comps, patterns, shifts, monthly);
    write_overview(overview_dir / ("spoor1_overzicht_dpt_" + D + ".xlsx"), summary, roster, comps, patterns, shifts, req);
    write_readable_roster(overview_dir / ("Cyclic_roster_dpt_" + D + "_leesbaar.txt"), dept, roster, comps, patterns, shifts);
    return summary;
}

fs::path project_root(const char* argv0) {
    // De projectmap: de eerste map met een submap data/, gezocht vanaf de map waar je het
    // programma start en daarna vanaf de map van spoor1.exe (telkens tot 3 mappen hoger).
    vector<fs::path> starts = {fs::current_path()};
    std::error_code ec;
    fs::path exe = fs::absolute(fs::path(argv0), ec).parent_path();
    if (!ec) starts.push_back(exe);
    for (fs::path d : starts) {
        for (int k = 0; k < 4; ++k) {
            if (fs::is_directory(d / DATA_DIR, ec)) return d;
            d = d.parent_path();
        }
    }
    throw ExitError("Geen map '" + DATA_DIR + "/' gevonden boven " + fs::current_path().string() + ".\n"
                    "Start spoor1.exe vanuit de projectmap (de map met data/, files/, results/).");
}

int run(int argc, char** argv) {
    fs::path root = project_root(argv[0]);
    string excel = (root / DATA_DIR / EXCEL_FILE).string();
    string depts_arg = DEPARTMENTS;
    string out = (root / FILES_DIR).string();
    string results_dir = (root / RESULTS_DIR).string();
    for (int k = 1; k < argc; ++k) {                            // --excel X, --dept A, --out X, --results X
        string a = argv[k], val;
        if (a == "-h" || a == "--help") {
            std::cout << "gebruik: spoor1.exe [--excel EXCEL] [--dept DEPT] [--out OUT] [--results RESULTS]\n\n"
                         "Case E spoor 1: cyclisch rooster met namen\n";
            return 0;
        }
        size_t eq = a.find('=');
        if (eq != string::npos) { val = a.substr(eq + 1); a = a.substr(0, eq); }
        else if (k + 1 < argc) val = argv[++k];
        else throw ExitError("argument " + a + ": verwacht een waarde");
        if (a == "--excel") excel = val;
        else if (a == "--dept") depts_arg = val;
        else if (a == "--out") out = val;
        else if (a == "--results") results_dir = val;
        else throw ExitError("onbekend argument: " + a);
    }

    if (!fs::exists(fs::path(excel)))
        throw ExitError("Excel niet gevonden: " + excel + "\nZet '" + EXCEL_FILE + "' in de map " + DATA_DIR +
                        "/ (of pas EXCEL_FILE bovenaan aan).");
    Workbook wb{fs::path(excel)};
    if (!wb.has_sheet(SHEET_CASE_D))
        throw ExitError("Tabblad '" + SHEET_CASE_D + "' ontbreekt in " + excel + ". Is dit de juiste versie van de Excel?");
    fs::create_directories(fs::path(results_dir));

    vector<std::pair<char, Summary>> results;
    for (char d : depts_arg) {
        char D = char(std::toupper((unsigned char)d));
        Summary s = run_department(wb, D, fs::path(out), fs::path(results_dir));
        auto it = std::find_if(results.begin(), results.end(), [&](const auto& x) { return x.first == D; });
        if (it == results.end()) results.emplace_back(D, s);
        else it->second = s;
    }
    if (results.empty()) throw std::runtime_error("geen afdelingen opgegeven");

    size_t w = 0;                                               // kolombreedte
    for (const auto& r : results)
        for (const auto& kv : r.second) w = std::max(w, u8len(kv.second.str()));
    w += 2;
    std::cout << "\nSpoor 1 (" << SCENARIO << ", " << POLICY << ")\n" << string(58 + w * results.size(), '-') << "\n";
    string line = pad_right("", 58);
    for (const auto& r : results) line += pad_left(string(1, r.first), w);
    std::cout << line << "\n";
    for (size_t k = 1; k < results[0].second.size(); ++k) {
        line = pad_right(results[0].second[k].first, 58);
        for (const auto& r : results) line += pad_left(r.second[k].second.str(), w);
        std::cout << line << "\n";
    }
    std::cout << "\nExcel gelezen uit:         " << excel << "\n";
    std::cout << "Invoer voor roster.exe in: " << out << "\n";
    std::cout << "Overzicht-Excels in:       " << results_dir << "\n";
    return 0;
}

int main(int argc, char** argv) {
    try {
        return run(argc, argv);
    } catch (const ExitError& e) {
        std::cerr << e.what() << "\n";
        return 1;
    } catch (const std::exception& e) {
        std::cerr << "FOUT: " << e.what() << "\n";
        return 1;
    }
}