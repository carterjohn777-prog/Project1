// ======================================================================
// Buffer tables + enums used by codeunit 50100 "GL Subledger Recon Mgt."
//
// Compare these against your existing objects -- the two things that
// matter are (a) explicit Integer primary keys with NO AutoIncrement
// (AutoIncrement is ignored on temporary records, which is what produced
// the "Entry No.='1' already exists" error), and (b) the secondary keys
// the Excel writers call with SetCurrentKey.
// ======================================================================

enum 50100 "Recon Account Category"
{
    Extensible = true;

    value(0; " ") { Caption = ' '; }
    value(1; Revenue) { Caption = 'Revenue'; }
    value(2; Expense) { Caption = 'Expense'; }
}

enum 50101 "Recon Party Type"
{
    Extensible = true;

    value(0; " ") { Caption = ' '; }
    value(1; Vendor) { Caption = 'Vendor'; }
    value(2; Customer) { Caption = 'Customer'; }
}

table 50100 "Recon GL Buffer"
{
    Caption = 'Recon GL Buffer';
    TableType = Temporary;
    DataClassification = SystemMetadata;

    fields
    {
        field(1; "Entry No."; Integer)
        {
            Caption = 'Entry No.';
            // No AutoIncrement -- assigned by the codeunit.
        }
        field(10; "Account No."; Code[20]) { Caption = 'Account No.'; }
        field(11; "Account Name"; Text[100]) { Caption = 'Account Name'; }
        field(12; "Account Category"; Enum "Recon Account Category") { Caption = 'Account Category'; }
        field(20; "Period Start Date"; Date) { Caption = 'Period Start Date'; }
        field(21; "Period Label"; Text[10]) { Caption = 'Period Label'; }
        field(30; "GL Amount"; Decimal) { Caption = 'GL Amount'; }
        field(31; "Subledger Amount"; Decimal) { Caption = 'Subledger Amount'; }
        field(32; "Without Vendor Amount"; Decimal) { Caption = 'Without Vendor Amount'; }
        field(33; "Without Customer Amount"; Decimal) { Caption = 'Without Customer Amount'; }
        field(34; Variance; Decimal) { Caption = 'Variance'; }
    }

    keys
    {
        key(PK; "Entry No.") { Clustered = true; }
        key(AccountPeriod; "Account No.", "Period Start Date") { }
        key(PeriodAccount; "Period Start Date", "Account No.") { }
    }
}

table 50101 "Recon Party Buffer"
{
    Caption = 'Recon Party Buffer';
    TableType = Temporary;
    DataClassification = SystemMetadata;

    fields
    {
        field(1; "Party Type"; Enum "Recon Party Type") { Caption = 'Party Type'; }
        field(2; "Party No."; Code[20]) { Caption = 'Party No.'; }
        field(3; "Period Start Date"; Date) { Caption = 'Period Start Date'; }
        field(10; "Party Name"; Text[100]) { Caption = 'Party Name'; }
        field(11; "Period Label"; Text[10]) { Caption = 'Period Label'; }
        field(20; Amount; Decimal) { Caption = 'Amount'; }
    }

    keys
    {
        // PK order matches SetCurrentKey in WritePartySheet, so the
        // single-pass grouping there is guaranteed to be contiguous.
        key(PK; "Party Type", "Party No.", "Period Start Date") { Clustered = true; }
    }
}

table 50102 "Recon Flat Buffer"
{
    Caption = 'Recon Flat Buffer';
    TableType = Temporary;
    DataClassification = SystemMetadata;

    fields
    {
        field(1; "Entry No."; Integer)
        {
            Caption = 'Entry No.';
            // No AutoIncrement -- assigned by the codeunit.
        }
        field(10; "Period Start Date"; Date) { Caption = 'Period Start Date'; }
        field(11; "Period Label"; Text[10]) { Caption = 'Period Label'; }
        field(20; "Account No."; Code[20]) { Caption = 'Account No.'; }
        field(21; "Account Name"; Text[100]) { Caption = 'Account Name'; }
        field(22; "Account Category"; Enum "Recon Account Category") { Caption = 'Account Category'; }
        field(30; Source; Text[50]) { Caption = 'Source'; }
        field(31; "Party Type"; Enum "Recon Party Type") { Caption = 'Party Type'; }
        field(32; "Party No."; Code[20]) { Caption = 'Party No.'; }
        field(33; "Party Name"; Text[100]) { Caption = 'Party Name'; }
        field(34; "Is Subledger"; Boolean) { Caption = 'Is Subledger'; }
        field(40; "Document No."; Code[20]) { Caption = 'Document No.'; }
        field(41; "G/L Entry No."; Integer) { Caption = 'G/L Entry No.'; }
        field(42; "Posting Date"; Date) { Caption = 'Posting Date'; }
        field(50; Amount; Decimal) { Caption = 'Amount'; }
    }

    keys
    {
        key(PK; "Entry No.") { Clustered = true; }
        key(PeriodAccount; "Period Start Date", "Account No.") { }
        key(Party; "Party Type", "Party No.") { }
    }
}
