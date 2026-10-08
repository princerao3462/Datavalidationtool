# Header Combinations Used by Data Validation Tool v12 — Generalized

This is the hybrid header recognition layer added to v12.

## Employee ID
Employee ID; EmployeeID; Employee Code; Employee No; Employee Number; Emp ID; EmpID; Emp Code; Emp No; Emp Number; Member ID; MemberID; Member Code; Staff ID; Staff No; Staff Number; Associate ID; Worker ID; Insured ID; ID.

## Employee Name
Employee Name; Employee Full Name; Employee Fullname; Emp Name; Emp Full Name; Emp Nm; Employee Nm; Member Name; Member Full Name; Insured Name; Insured Full Name; Self Name; Self Full Name; Name of Employee; Name of Insured; Employee; Member; Insured.

## DOB
Date of Birth; Date Birth; Birth Date; Birthdate; Birth Dt; DOB; D O B; Dateofbirth.

## Age
Age; Employee Age; Member Age; Insured Age; Self Age; Age in Years; Years Age.

## Gender
Gender; Sex; Employee Gender; Member Gender; Insured Gender; Self Gender.

## Relationship
Relationship; Relation; Relationship to Self; Relation to Self; Relationship with Employee; Relation with Employee; Relation 2; Relation Type; Dependent Type; Dependent Relationship; Dep Type; Member Relation.

## Phone
Mobile Number; Mobile No; Mobile; Mobile Phone; Phone Number; Phone No; Phone; Telephone; Contact Number; Contact No; Contact; Cell; Employee Mobile; Member Mobile.

## GHI / Sum Insured
GHI Sum Insured; Sum Insured; Sum Insured Amount; GHI; GHI Amount; Insurance Amount; Insured Amount; Coverage Amount; Coverage; SI Amount; SI.

## Wide-format groups

### Employee / Self
Employee Name, Employee Full Name, Emp Name, Employee DOB, Employee Age, Employee Gender, etc.

### Spouse
Spouse Name, Spouse DOB, Spouse Age, Wife Name, Husband Name, etc.

### Children
Child 1 Name / DOB / Age / Gender / Mobile; Child 2 ...; Child 3 ...; and so on. Child number is dynamic and is not tied to a column position.

### Parents
Father Name / DOB / Age / Gender; Mother Name / DOB / Age / Gender; Parent Name / DOB / Age. These map to Parent.

### Parents-in-law
Father-in-law Name / DOB / Age / Gender; Mother-in-law Name / DOB / Age / Gender; Father in Law ...; Mother in Law ...; Parent in Law ... . These map to Parent-in-law.

## Typo handling

The tool uses conservative fuzzy matching for simple spelling/spacing/punctuation mistakes. Examples that can be recognized when sufficiently close:

- Employe ID → Employee ID
- Employee Nam → Employee Name
- Employee Ful Name → Employee Full Name
- Date of Brith → Date of Birth
- Gendr → Gender
- Relatonship → Relationship

It deliberately does **not** make low-confidence guesses such as `E-ID` unless the surrounding header/value context makes the meaning clear.

## User fallback

If a header is too complex, rename it to one of the standard headers shown in the tool. Examples:

- E-ID → Employee ID
- EMP Nm → Employee Name
- DOB Dt → Employee DOB (if this is the employee DOB)
- Dep Type → Relationship

Column order is never part of the rule.


## v3 fix for this source structure

The following structure is explicitly supported:

Employee Name
Employee Date of Birth
Employee Age
Marital Status
Spouse Name
Date of Birth
Number of Children
Child 1 - Name
Date of Birth
Child 2 - Name
Date of Birth
Child 3-Name
Date of Birth
Father's Name
Father's Age
Mother's Name
Mother's Age
Mobile Number

Repeated generic `Date of Birth` columns are associated contextually with
the nearest preceding member group/name header and bounded by the next
member name header. This means the tool can distinguish Spouse DOB,
Child 1 DOB, Child 2 DOB, Child 3 DOB and Employee DOB without fixed
column positions.


## Regression case now supported

The classifier and wide parser explicitly support:

Employee Name
Employee Date of Birth
Employee Age
Marital Status
Spouse Name
Date of Birth
Number of Children
Child 1 - Name
Date of Birth
Child 2 - Name
Date of Birth
Child 3-Name
Date of Birth
Father's Name
Father's Age
Mother's Name
Mother's Age
Mobile Number

Generic repeated `Date of Birth` headers are interpreted according to the
nearest preceding member-name header. No Excel column letter or fixed
position is used.


## Definitive wide-format fallback
The processing pipeline now runs a semantic wide-format override immediately
after legacy classification and again at the rejection gate. Therefore a file
with Employee Name + Employee Date of Birth + Spouse/Child/Father/Mother
headers is routed to the wide processor even if the legacy classifier returns
unknown. This fallback does not use column letters or client names.

## Standard Relationship Codes
The engine also recognizes these relationship/member values:
`MMBR`, `Member`, `SPSE`, `UDTR`, `Daughter`, `SONM`, `Son`, `EMOT`,
`EFAT`, `EFLW`, and `EMLW`. These are semantic relationship aliases, not
client-specific rules.

## Parent-in-law relationship precedence
Any relationship value containing a clear `law` / `in-law` / `inlaw`
relationship indicator is mapped to `Parent-in-law`. This includes EFLW,
EMLW, M law, F law, mother-law, father-law, mother-in-law, father-in-law,
parent law, and equivalent spellings.
