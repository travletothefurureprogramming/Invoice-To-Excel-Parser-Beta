# 📄 Invoice2Excel Pro Beta Edition

> **Automated PDF Invoice Data Extraction Tool**
>
> Ένα τοπικό, γρήγορο και ασφαλές εργαλείο για αυτόματη ανάγνωση PDF τιμολογίων και εξαγωγή των βασικών οικονομικών στοιχείων σε Excel.

![Status](https://img.shields.io/badge/status-beta-orange)
![Version](https://img.shields.io/badge/version-1.0.0-blue)
![Python](https://img.shields.io/badge/python-3.9%2B-blue)
![License](https://img.shields.io/badge/license-GPL--3.0-green)

---

## ⚡ Σχετικά με το Project

Το **Invoice2Excel Pro** είναι μια desktop εφαρμογή γραμμένη σε **Python**, με γραφικό περιβάλλον βασισμένο στο **CustomTkinter**.

Σχεδιάστηκε για να αυτοματοποιεί την επεξεργασία φακέλων με PDF τιμολόγια. Η εφαρμογή διαβάζει το text layer των PDF, εντοπίζει σημαντικές πληροφορίες και δημιουργεί ένα οργανωμένο αρχείο **`.xlsx`**.

Για κάθε τιμολόγιο μπορεί να εξάγει:

* 📅 Ημερομηνία τιμολογίου
* 🧾 ΑΦΜ / Tax ID / VAT
* 💰 Καθαρή αξία
* 🧮 Ποσό ΦΠΑ
* 💵 Συνολικό ποσό
* 📁 Όνομα αρχείου
* ⚠️ Κατάσταση επεξεργασίας

Τα παραπάνω αποθηκεύονται σε δομημένες στήλες στο Excel.

---

## ✨ Κύρια Χαρακτηριστικά

### 🧾 Έξυπνη εξαγωγή Tax ID / VAT

Το parser χρησιμοποιεί πολλαπλά επίπεδα αναγνώρισης:

1. Ελληνικό **ΑΦΜ**
2. Αμερικανικό **EIN**
3. EU / UK VAT με country prefix
4. Fallback αναζήτηση EU / UK VAT
5. Fallback για standalone 9-digit Tax ID

Υποστηρίζονται VAT prefixes από πολλές χώρες της ΕΕ και το Ηνωμένο Βασίλειο.

Πριν από την αναζήτηση Tax ID αφαιρούνται επίσης email addresses και URLs, ώστε να μειώνονται false positives.

---

### 📅 Αναγνώριση ημερομηνιών

Ο parser υποστηρίζει διαφορετικές μορφές ημερομηνίας, όπως:

```text
2026-01-15
15/01/2026
15.01.2026
15-01-26
01/15/2026
15 January 2026
January 15, 2026
15 Ιανουαρίου 2026
15 Ιαν 2026
```

Υποστηρίζονται ονόματα μηνών στα:

* 🇬🇷 Ελληνικά
* 🇬🇧 Αγγλικά
* 🇩🇪 Γερμανικά
* 🇫🇷 Γαλλικά
* 🇪🇸 Ισπανικά

Οι ημερομηνίες μετατρέπονται σε μορφή **ISO `YYYY-MM-DD`**.

---

### 💰 Αναγνώριση ποσών

Το Invoice2Excel Pro αναγνωρίζει διαφορετικές μορφές οικονομικών ποσών, συμπεριλαμβανομένων ευρωπαϊκών και αμερικανικών separators.

Παραδείγματα:

```text
1.234,56 €
€1,234.56
1234.56 EUR
1 234,56 €
$1,250.00
1,250.00 USD
```

Υποστηρίζονται μεταξύ άλλων:

* EUR / €
* USD / $
* GBP / £
* CHF
* CNY / ¥
* BRL / R$

Το parser προσπαθεί να αναγνωρίσει ξεχωριστά:

* **Net Amount**
* **VAT Amount**
* **Total Amount**

μέσω σχετικών keywords σε διαφορετικές γλώσσες.

---

### 🧮 Έξυπνοι υπολογισμοί ποσών

Όταν κάποια οικονομική πληροφορία λείπει, ο parser μπορεί να την υπολογίσει από τα διαθέσιμα δεδομένα.

Για παράδειγμα:

```text
Net = Total - VAT
```

ή:

```text
Total = Net + VAT
```

Εάν υπάρχει καθαρή αξία και γνωστό VAT rate, μπορεί επίσης να υπολογιστεί:

```text
VAT = Net × VAT Rate / 100
```

Σε περίπτωση που δεν υπάρχει keyword για το Total, χρησιμοποιείται ως fallback το μεγαλύτερο invoice-like ποσό που εντοπίστηκε. Οι υπολογισμοί συνοδεύονται από warnings ώστε ο χρήστης να γνωρίζει ότι η τιμή δεν βρέθηκε απευθείας.

---

## 🖥️ Desktop GUI

Η εφαρμογή διαθέτει γραφικό περιβάλλον με:

* 📂 Επιλογή φακέλου εισόδου
* 📊 Επιλογή αρχείου Excel εξόδου
* ▶️ Start Processing
* 📈 Progress bar
* 📝 Real-time processing log
* ⚠️ Error / warning reporting
* 🌙 Dark interface

Το GUI χρησιμοποιεί background worker thread ώστε η επεξεργασία να μην μπλοκάρει το κύριο interface.

---

## 📊 Excel Export

Το αποτέλεσμα αποθηκεύεται σε αρχείο **`.xlsx`** με worksheet:

```text
Invoices
```

Οι στήλες είναι:

| Column               | Περιγραφή             |
| -------------------- | --------------------- |
| `Invoice Date`       | Ημερομηνία τιμολογίου |
| `Tax ID / VAT`       | ΑΦΜ / VAT / Tax ID    |
| `Net Amount (€/$)`   | Καθαρή αξία           |
| `VAT Amount (€/$)`   | Ποσό ΦΠΑ              |
| `Total Amount (€/$)` | Συνολικό ποσό         |
| `File Name`          | Όνομα PDF             |
| `Status`             | Success / Warning     |

Το Excel περιλαμβάνει επιπλέον:

* Freeze panes
* Auto-filter
* Automatic column sizing
* Number formatting
* Borders
* Formatted headers
* Status highlighting
* Απενεργοποιημένα gridlines

Η εφαρμογή χρησιμοποιεί προσωρινό αρχείο και atomic replacement όπου υποστηρίζεται, μειώνοντας τον κίνδυνο incomplete output files.

---

# 🔒 Privacy & Local Processing

Το Invoice2Excel Pro έχει σχεδιαστεί για **τοπική επεξεργασία**.

Η αρχιτεκτονική του parser ανοίγει τα PDF μέσω `pdfplumber`, εξάγει το text layer και επεξεργάζεται τα δεδομένα τοπικά.

### Τι σημαίνει αυτό;

* 📌 Τα PDF επεξεργάζονται στον υπολογιστή σας.
* 📌 Δεν απαιτείται cloud service για το parsing.
* 📌 Δεν απαιτείται API key.
* 📌 Δεν απαιτείται online AI service.
* 📌 Τα δεδομένα που εξάγονται παραμένουν τοπικά, εκτός εάν ο χρήστης τα μεταφέρει ο ίδιος αλλού.

> **Privacy note:** Η εφαρμογή είναι σχεδιασμένη για local processing και δεν περιλαμβάνει μηχανισμό αποστολής των invoice δεδομένων σε cloud API. Η φράση "GDPR compliant" δεν αποτελεί από μόνη της νομική πιστοποίηση συμμόρφωσης· η συνολική συμμόρφωση εξαρτάται και από τον τρόπο με τον οποίο ο χρήστης/οργανισμός διαχειρίζεται τα δεδομένα.

---

# ⚠️ Known Limitations

## 1. Scanned / Image-only PDFs

Η εφαρμογή **δεν περιλαμβάνει OCR engine**.

Εάν ένα PDF αποτελείται αποκλειστικά από εικόνες, το `pdfplumber` μπορεί να μην επιστρέψει text layer.

Σε αυτή την περίπτωση εμφανίζεται warning όπως:

```text
No text layer was extracted; scanned/image-only PDF requires OCR.
```

### Παράδειγμα

❌ Scanned paper invoice:

```text
[IMAGE OF INVOICE]
```

δεν υποστηρίζεται χωρίς εξωτερικό OCR.

✅ Digital PDF:

```text
Invoice Date: 15/08/2026
VAT: EL123456789
Total: 124.00 €
```

μπορεί να αναλυθεί κανονικά.

---

## 2. Μη τυπικά Invoice Layouts

Το parser βασίζεται σε:

* text extraction
* keywords
* regular expressions
* structural rules
* fallback logic

Επομένως, εξαιρετικά ασυνήθιστα layouts ή invoices με πολύ διαφορετική δομή ενδέχεται να χρειάζονται χειροκίνητο έλεγχο.

---

## 3. Extraction ≠ Accounting Verification

Το πρόγραμμα **δεν αποτελεί λογιστικό ή φορολογικό σύστημα**.

Το αποτέλεσμα πρέπει να ελέγχεται πριν χρησιμοποιηθεί για:

* φορολογικές δηλώσεις
* λογιστικές εγγραφές
* επίσημα οικονομικά αρχεία
* πληρωμές
* οικονομικές αποφάσεις

---

# 🚀 Installation

## Requirements

* Python **3.9 ή νεότερη**
* Windows / Linux / macOS
* PDF files με readable text layer

### Python dependencies

```text
customtkinter
pandas
pdfplumber
openpyxl
```

---

## 1. Clone το Repository

```bash
git clone https://github.com/travletothefurureprogramming/Invoice-To-Excel-Parser-Beta.git
cd Invoice-To-Excel-Parser-Beta
```

---

## 2. Δημιουργία Virtual Environment

### Windows

```powershell
python -m venv .venv
.venv\Scripts\activate
```

### Linux / macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
```

---

## 3. Εγκατάσταση Dependencies

```bash
pip install -r requirements.txt
```

Εναλλακτικά:

```bash
pip install customtkinter pandas pdfplumber openpyxl
```

---

## 4. Εκκίνηση

```bash
python main.py
```

---

# 🧑‍💻 Usage

### Βήμα 1

Επίλεξε τον φάκελο που περιέχει τα PDF invoices.

```text
Input PDF Folder
```

### Βήμα 2

Επίλεξε το αρχείο Excel που θέλεις να δημιουργηθεί.

```text
Output XLSX
```

### Βήμα 3

Πάτησε:

```text
Start Processing
```

Η εφαρμογή εντοπίζει όλα τα `.pdf` αρχεία στον επιλεγμένο φάκελο και τα επεξεργάζεται με αλφαβητική σειρά.

### Βήμα 4

Παρακολούθησε το progress bar και το log.

Παράδειγμα:

```text
Found 10 PDF file(s).

Processing file 1 of 10: invoice_001.pdf
[OK] invoice_001.pdf

Processing file 2 of 10: invoice_002.pdf
[WARN] invoice_002.pdf: Invoice Date not found.

...

[DONE] Excel export created
Finished. Processed 10 file(s).
```

---

# 📁 Project Structure

```text
PDF-Invoice-Parser/
│
├── main.py
├── requirements.txt
├── LICENSE
└── README.md
```

### `main.py`

Κύριος κώδικας της εφαρμογής:

* GUI
* PDF text extraction
* Tax ID / VAT extraction
* Date parsing
* Amount extraction
* Validation
* Warning handling
* Excel export
* Background processing

### `requirements.txt`

Python dependencies του project.

### `LICENSE`

Πλήρες κείμενο της GNU GPL v3.0.

### `README.md`

Documentation του project.

---

# 🏗️ Architecture

Η βασική ροή του προγράμματος είναι:

```text
             PDF Folder
                  │
                  ▼
          ┌───────────────┐
          │   pdfplumber  │
          └───────┬───────┘
                  │
                  ▼
            Text Layer
                  │
       ┌──────────┼──────────┐
       ▼          ▼          ▼
     Date       Tax ID     Amounts
       │          │          │
       └──────────┼──────────┘
                  ▼
            ParseResult
                  │
                  ▼
          Validation / Warnings
                  │
                  ▼
            pandas DataFrame
                  │
                  ▼
             openpyxl
                  │
                  ▼
              XLSX File
```

Η επεξεργασία πραγματοποιείται σε background worker thread, ενώ το GUI λαμβάνει progress/status messages μέσω queue.

---

# 🧪 Beta Status

Το **Invoice2Excel Pro 1.0.0** βρίσκεται σε **Beta**.

Αυτό σημαίνει ότι:

* ενδέχεται να υπάρχουν bugs
* ορισμένα invoice layouts μπορεί να μην υποστηρίζονται
* τα extracted δεδομένα πρέπει να ελέγχονται
* το OCR δεν υποστηρίζεται ακόμη
* η εφαρμογή δεν πρέπει να θεωρείται υποκατάστατο λογιστικού λογισμικού

---

# ⚖️ Disclaimer

> **DISCLAIMER OF WARRANTY & LIMITATION OF LIABILITY**
>
> Το Invoice2Excel Pro παρέχεται **"AS IS"**, χωρίς εγγύηση ότι τα αποτελέσματα extraction είναι πάντοτε πλήρη, ακριβή ή κατάλληλα για συγκεκριμένο σκοπό.
>
> Ο χρήστης είναι υπεύθυνος για την επαλήθευση των δεδομένων που παράγονται από την εφαρμογή πριν από οποιαδήποτε λογιστική, φορολογική, οικονομική ή άλλη επίσημη χρήση.
>
> Ο δημιουργός δεν ευθύνεται για τυχόν λάθη στην εξαγωγή δεδομένων, λανθασμένες καταχωρήσεις, απώλεια δεδομένων ή οικονομικές/φορολογικές συνέπειες που μπορεί να προκύψουν από τη χρήση του λογισμικού.
>
> Για τους πλήρεις όρους warranty disclaimer και limitation of liability ισχύει το κείμενο της **GNU GPL v3.0** που περιλαμβάνεται στο αρχείο `LICENSE`.

---

## 🤝 Feedback, Bug Reports & Roadmap

Για τη διατήρηση της πλήρους κυριότητας και των δικαιωμάτων πνευματικής ιδιοκτησίας (copyright) του έργου, **δεν γίνονται δεκτές υποβολές κώδικα (Pull Requests / Code Contributions) από τρίτους**.

Ωστόσο, η υποστήριξη και η γνώμη σας είναι πολύτιμη! Μπορείτε να συνεισφέρετε ανοίγοντας **Issues**:
* 🐛 **Bug Reports:** Αναφορά σφαλμάτων ή δυσλειτουργιών.
* 💡 **Feature Requests:** Προτάσεις για νέα χαρακτηριστικά και βελτιώσεις.

### 🎯 Προγραμματισμένες Βελτιώσεις (Roadmap από τον δημιουργό)
* 🔍 **OCR Support:** Υποστήριξη Tesseract/EasyOCR για σκαναρισμένα PDF.
* 🧠 **Layout Expansion:** Υποστήριξη περισσότερων layouts τιμολογίων & αποδείξεων.
* 🌍 **International Formats:** Υποστήριξη διεθνών formats και πολλαπλών νομισμάτων.
* 🧾 **More Invoice Fields:** Αναγνώριση επιπλέον πεδίων (ΦΠΑ, καθαρή αξία, περιγραφή γραμμών).
* 📊 **Advanced Excel Reporting:** Προσαρμοσμένα πρότυπα αναφορών στο Excel.
* ⚡ **Parallel Processing:** Παράλληλη επεξεργασία αρχείων για ακόμα μεγαλύτερη ταχύτητα.
* 🖥️ **Executable Releases:** Πακέτα εγκατάστασης (`.exe` / `.dmg`) για εύκολη χρήση χωρίς Python.

---

# 📜 License

Το **Invoice2Excel Pro** διανέμεται υπό την άδεια:

**GNU General Public License v3.0 (GPL-3.0)**

Η GPL-3.0 επιτρέπει:

* ✅ Χρήση
* ✅ Μελέτη
* ✅ Τροποποίηση
* ✅ Αναδιανομή
* ✅ Δημιουργία παράγωγων έργων

Υπό τους όρους της GPL-3.0, οι σχετικές απαιτήσεις copyleft και source-code distribution ισχύουν για έργα που διανέμονται σύμφωνα με την άδεια.

Για το πλήρες νομικό κείμενο, δείτε:

```text
LICENSE
```

---

## 🔒 Απόρρητο & Όροι Χρήσης

Το Invoice2Excel επεξεργάζεται τα αρχεία τιμολογίων **τοπικά στον υπολογιστή σας**.

Τα δεδομένα των τιμολογίων **δεν ανεβαίνουν στους διακομιστές μας και δεν αποθηκεύονται από εμάς**.

Με τη λήψη ή τη χρήση του Invoice2Excel, αποδέχεστε την ισχύουσα
[Πολιτική Απορρήτου και τους Όρους Χρήσης](https://docs.google.com/document/d/1MF9LJrt18siQNvx-H1w_GDlq-wuklhJhd5z8pciY2Rc/edit?usp=sharing) 


> **⚠️ Έκδοση Beta:** Το Invoice2Excel βρίσκεται αυτή τη στιγμή σε έκδοση Beta.
> Πριν χρησιμοποιήσετε τα εξαγόμενα δεδομένα για λογιστικούς, φορολογικούς,
> οικονομικούς ή άλλους επίσημους σκοπούς, **ελέγχετε πάντα τα αποτελέσματα
> σε σύγκριση με το αρχικό τιμολόγιο**.

---

## 🎨 Icon Attribution

The application icon **“Rescan Document”** is provided by **Icons8**.

* Icon: [Rescan Document](https://icons8.com/icon/84iQU-5TSpyl/rescan-document)
* Provider: [Icons8](https://icons8.com)

The icon is used under the applicable Icons8 licensing and attribution terms.

---

# ⭐ Project

**Invoice2Excel Pro**

> Turn invoices into structured data.
> **Locally. Automatically. Simply.**
