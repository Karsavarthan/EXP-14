# Exercise-14-Capstone-Project
# Automated Certificate Generation using UiPath
## Aim
Automatically generate personalized certificates for participants by reading a CSV file and filling a certificate template (Word or PowerPoint), then saving each certificate as PDF (and optionally emailing or zipping them).

## Materials Required
1) UiPath Studio (Community/Enterprise)
2) Windows OS
3) Certificate template (Word .docx or PowerPoint .pptx) with clear placeholders (e.g., {{Name}}, {{Course}}, {{Date}}, {{Grade}})
4) Input CSV file (participants.csv)
5) Basic UiPath knowledge:
6) Read CSV / Read Range
7) Word/PowerPoint activities (Word Application Scope / PowerPoint activities or use “Word Application Scope” + “Replace Text”)
8) Save As / Export to PDF / Save Presentation As
9) Send SMTP Mail Message or Outlook Mail Message (optional)
10) Invoke VBA (optional)
11) System.IO for file operations

## Input (CSV example)

Create participants.csv with columns:
```
Name,Course,Date,Grade,Email
Asha Kumar,Deep Learning Basics,24-11-2025,Distinction,asha@example.com
Ramesh Iyer,Deep Learning Basics,24-11-2025,Merit,ramesh@example.com
```

## Template Requirements

Word template (certificate_template.docx) or PowerPoint template (certificate_template.pptx)

Placeholders must be unique and easy to locate. Use double braces, e.g. {{Name}}, {{Course}}, {{Date}}, {{Grade}}.

For Word: placeholders in normal text or content controls. For PowerPoint: placeholders inside textboxes.

## Overall Workflow (high-level)

Read CSV → dtParticipants

For Each row In dtParticipants

Set variables: name, course, date, grade, email

Copy template to a working file (e.g., temp_certificate_{Name}.docx / .pptx)

Open the working file and replace placeholders with values

Save / Export working file as PDF (certificate_{Name}.pdf)

(Optional) Attach & send email or move to folder / zip

End

## Detailed UiPath Sequence (Word-based preferred approach)
# Variables (suggested)

dtParticipants — DataTable

row — DataRow

name — String

course — String

certDate — String

grade — String

email — String

templatePath — String → "C:\Templates\certificate_template.docx"

outputFolder — String → "C:\Certificates\"

tempDocPath — String

pdfPath — String

# Sequence :

<img width="899" height="497" alt="image" src="https://github.com/user-attachments/assets/d52c0489-b1f3-41aa-a851-eea82493da30" />

<img width="843" height="564" alt="image" src="https://github.com/user-attachments/assets/6fae79c8-a2de-4f63-99ff-d3a8f1085d03" />

<img width="875" height="652" alt="image" src="https://github.com/user-attachments/assets/9200a2bd-e222-4bb3-a298-56ebf0e18214" />


# Input:

<img width="1912" height="1063" alt="image" src="https://github.com/user-attachments/assets/6be682cc-548c-4ee4-85ee-3acbf6dad7f7" />


# Output :

<img width="922" height="710" alt="image" src="https://github.com/user-attachments/assets/df2c725c-9ef4-4e06-bd67-d8f56059d20c" />

# Result :

The UiPath workflow successfully reads the list of participants from a CSV file, replaces the template placeholders with personalized details, and automatically generates individual certificates for each participant in PDF/Word format. This automation eliminates manual editing work, ensures accuracy and consistency, and produces certificates efficiently for any number of students.
