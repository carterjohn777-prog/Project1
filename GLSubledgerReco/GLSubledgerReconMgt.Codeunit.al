codeunit 50100 "GL Subledger Recon Mgt."
{
    // ==================================================================
    // ENTRY POINT
    // ==================================================================
    // AsOfDate: any date inside the last month you want included. The
    // report covers that month plus the 5 preceding calendar months.
    //
    // UseStrictLineValidation: when TRUE a G/L Entry only counts as
    // "subledger" if the posted document also carries a line of
    // Type = G/L Account for that exact account. Leave FALSE for most
    // companies -- Item/Resource lines post to Sales/COGS via Posting
    // Setup, not via a G/L Account line.
    procedure RunExport(AsOfDate: Date; UseStrictLineValidation: Boolean)
    var
        MonthStartDate: array[6] of Date;
        MonthLabel: array[6] of Text[10];
    begin
        StrictLineValidation := UseStrictLineValidation;

        // FIX 1: entry-number counters live in the codeunit, not in a
        // shadowed local record. See GetNextGLBufEntryNo below.
        NextGLBufEntryNo := 0;
        NextFlatBufEntryNo := 0;

        BuildMonthList(AsOfDate, MonthStartDate, MonthLabel);

        ReconGLBuffer.Reset();
        ReconGLBuffer.DeleteAll();
        ReconPartyBuffer.Reset();
        ReconPartyBuffer.DeleteAll();
        ReconFlatBuffer.Reset();
        ReconFlatBuffer.DeleteAll();

        PopulateBuffers(MonthStartDate, MonthLabel);
        ExportToExcel(MonthLabel);
    end;

    var
        ReconGLBuffer: Record "Recon GL Buffer" temporary;
        ReconPartyBuffer: Record "Recon Party Buffer" temporary;
        ReconFlatBuffer: Record "Recon Flat Buffer" temporary;
        StrictLineValidation: Boolean;
        NextGLBufEntryNo: Integer;
        NextFlatBufEntryNo: Integer;

    // ------------------------------------------------------------------
    // PERIOD SETUP
    // ------------------------------------------------------------------
    local procedure BuildMonthList(AsOfDate: Date; var MonthStartDate: array[6] of Date; var MonthLabel: array[6] of Text[10])
    var
        i: Integer;
        d: Date;
    begin
        d := CalcDate('<-CM-5M>', AsOfDate);
        for i := 1 to 6 do begin
            MonthStartDate[i] := d;
            MonthLabel[i] := Format(d, 0, '<Month Text,3>-<Year,2>'); // e.g. "Mar-26"
            d := CalcDate('<+1M>', d);
        end;
    end;

    local procedure MonthEndDate(MonthStart: Date): Date
    begin
        exit(CalcDate('<+1M-1D>', MonthStart));
    end;

    local procedure MonthIndex(MonthLabel: array[6] of Text[10]; PeriodLabel: Text[10]): Integer
    var
        i: Integer;
    begin
        for i := 1 to 6 do
            if MonthLabel[i] = PeriodLabel then
                exit(i);
        exit(0);
    end;

    // FIX 1: the original GetNextEntryNo() declared a LOCAL variable with
    // the same name as the global temporary buffer. The local one was a
    // NON-temporary record, so FindLast() read the (empty) physical table
    // and returned 1 on every call -- hence "The record in table Recon GL
    // Buffer already exists. Entry No.='1'". A simple counter removes the
    // dependency on the buffer's contents and its current filters.
    local procedure GetNextGLBufEntryNo(): Integer
    begin
        NextGLBufEntryNo += 1;
        exit(NextGLBufEntryNo);
    end;

    local procedure GetNextFlatBufEntryNo(): Integer
    begin
        NextFlatBufEntryNo += 1;
        exit(NextFlatBufEntryNo);
    end;

    // ------------------------------------------------------------------
    // MAIN LOOP: G/L Account x Month -> classify every G/L Entry
    // ------------------------------------------------------------------
    local procedure PopulateBuffers(MonthStartDate: array[6] of Date; MonthLabel: array[6] of Text[10])
    var
        GLAccount: Record "G/L Account";
        GLEntry: Record "G/L Entry";
        AccCategory: Enum "Recon Account Category";
        i: Integer;
        GLBufRecNo: Integer;
        GLAmt: Decimal;
        SubAmt: Decimal;
        WoVendorAmt: Decimal;
        WoCustomerAmt: Decimal;
    begin
        GLAccount.SetRange("Account Type", GLAccount."Account Type"::Posting);
        GLAccount.SetRange("Income/Balance", GLAccount."Income/Balance"::"Income Statement");
        GLAccount.SetFilter(
            "Account Category", '%1|%2|%3',
            GLAccount."Account Category"::Income,
            GLAccount."Account Category"::Expense,
            GLAccount."Account Category"::"Cost of Goods Sold");
        if GLAccount.FindSet() then
            repeat
                if GLAccount."Account Category" = GLAccount."Account Category"::Income then
                    AccCategory := AccCategory::Revenue
                else
                    AccCategory := AccCategory::Expense;

                for i := 1 to 6 do begin
                    GLBufRecNo := InsertGLBufferHeader(GLAccount, AccCategory, MonthStartDate[i], MonthLabel[i]);

                    // FIX 2: accumulate in local decimals and write the
                    // header row ONCE per Account x Month. The original
                    // did Get/Modify on the global buffer inside the entry
                    // loop, which is both slow and moves the record while
                    // other code may be relying on its position.
                    Clear(GLAmt);
                    Clear(SubAmt);
                    Clear(WoVendorAmt);
                    Clear(WoCustomerAmt);

                    GLEntry.Reset();
                    GLEntry.SetRange("G/L Account No.", GLAccount."No.");
                    GLEntry.SetRange("Posting Date", MonthStartDate[i], MonthEndDate(MonthStartDate[i]));
                    if GLEntry.FindSet() then
                        repeat
                            ClassifyAndAccumulate(
                                GLEntry, GLAccount, AccCategory, MonthStartDate[i], MonthLabel[i],
                                GLAmt, SubAmt, WoVendorAmt, WoCustomerAmt);
                        until GLEntry.Next() = 0;

                    UpdateGLBufferTotals(GLBufRecNo, GLAmt, SubAmt, WoVendorAmt, WoCustomerAmt);
                end;
            until GLAccount.Next() = 0;
    end;

    local procedure InsertGLBufferHeader(GLAccount: Record "G/L Account"; AccCategory: Enum "Recon Account Category"; PeriodStart: Date; PeriodLabel: Text[10]): Integer
    begin
        // One header row per Account x Month, created up front so the row
        // exists even for months with zero activity.
        ReconGLBuffer.Init();
        ReconGLBuffer."Entry No." := GetNextGLBufEntryNo();
        ReconGLBuffer."Account No." := GLAccount."No.";
        ReconGLBuffer."Account Name" := GLAccount.Name;
        ReconGLBuffer."Account Category" := AccCategory;
        ReconGLBuffer."Period Start Date" := PeriodStart;
        ReconGLBuffer."Period Label" := PeriodLabel;
        ReconGLBuffer.Insert(false);
        exit(ReconGLBuffer."Entry No.");
    end;

    local procedure UpdateGLBufferTotals(GLBufRecNo: Integer; GLAmt: Decimal; SubAmt: Decimal; WoVendorAmt: Decimal; WoCustomerAmt: Decimal)
    begin
        ReconGLBuffer.Reset();
        ReconGLBuffer.Get(GLBufRecNo);
        ReconGLBuffer."GL Amount" := GLAmt;
        ReconGLBuffer."Subledger Amount" := SubAmt;
        ReconGLBuffer."Without Vendor Amount" := WoVendorAmt;
        ReconGLBuffer."Without Customer Amount" := WoCustomerAmt;
        ReconGLBuffer.Variance := GLAmt - (SubAmt + WoVendorAmt + WoCustomerAmt);
        ReconGLBuffer.Modify(false);
    end;

    // ------------------------------------------------------------------
    // CLASSIFICATION: is this G/L Entry tied to a posted Vendor/Customer
    // document, or was it posted directly (journal / no subledger)?
    // ------------------------------------------------------------------
    local procedure ClassifyAndAccumulate(
        GLEntry: Record "G/L Entry";
        GLAccount: Record "G/L Account";
        AccCategory: Enum "Recon Account Category";
        PeriodStart: Date;
        PeriodLabel: Text[10];
        var GLAmt: Decimal;
        var SubAmt: Decimal;
        var WoVendorAmt: Decimal;
        var WoCustomerAmt: Decimal)
    var
        PartyNo: Code[20];
        PartyName: Text[100];
        PartyType: Enum "Recon Party Type";
        SourceText: Text[50];
        IsSubledger: Boolean;
        EntryAmount: Decimal;
    begin
        IsSubledger := false;
        PartyType := PartyType::" ";
        SourceText := '';

        // FIX 3: "Source Type" is blank on a fair number of G/L Entries
        // (posting groups, some posting routines, upgraded data). The
        // original CASE fell straight through to "Direct GL Journal" in
        // that situation. Blank now falls back to a document-number
        // lookup against all four posted headers before it is written off
        // as a direct journal.
        case GLEntry."Source Type" of
            GLEntry."Source Type"::Vendor:
                if TryFindVendorDocument(GLEntry, PartyNo, PartyName, SourceText) then
                    if PassesStrictCheck(GLEntry) then begin
                        IsSubledger := true;
                        PartyType := PartyType::Vendor;
                    end;
            GLEntry."Source Type"::Customer:
                if TryFindCustomerDocument(GLEntry, PartyNo, PartyName, SourceText) then
                    if PassesStrictCheck(GLEntry) then begin
                        IsSubledger := true;
                        PartyType := PartyType::Customer;
                    end;
            else begin
                if TryFindVendorDocument(GLEntry, PartyNo, PartyName, SourceText) then begin
                    IsSubledger := true;
                    PartyType := PartyType::Vendor;
                end else
                    if TryFindCustomerDocument(GLEntry, PartyNo, PartyName, SourceText) then begin
                        IsSubledger := true;
                        PartyType := PartyType::Customer;
                    end;
                if IsSubledger and not PassesStrictCheckForType(GLEntry, PartyType) then begin
                    IsSubledger := false;
                    PartyType := PartyType::" ";
                    PartyNo := '';
                    PartyName := '';
                    SourceText := '';
                end;
            end;
        end;

        if not IsSubledger then begin
            if AccCategory = AccCategory::Expense then
                SourceText := 'Direct GL Journal - No Vendor'
            else
                SourceText := 'Direct GL Journal - No Customer';
            PartyNo := '';
            PartyName := 'N/A - Direct Entry';
        end;

        // FIX 4: G/L Entry amounts are signed (revenue is a credit, i.e.
        // negative). Revenue is flipped to a positive presentation figure
        // so Revenue and Expense read the same way on the sheets. All four
        // buckets are flipped together, so the Variance tie-out is
        // unaffected. Delete NormalizeAmount() if you want raw GL signs.
        //GLEntry.CalcFields(Amount);
        EntryAmount := NormalizeAmount(AccCategory, GLEntry.Amount);

        GLAmt += EntryAmount;
        if IsSubledger then
            SubAmt += EntryAmount
        else
            if AccCategory = AccCategory::Expense then
                WoVendorAmt += EntryAmount
            else
                WoCustomerAmt += EntryAmount;

        // --- Vendor/Customer monthly buffer (Tab 3) ---
        if IsSubledger then
            AccumulatePartyBuffer(PartyType, PartyNo, PartyName, PeriodStart, PeriodLabel, EntryAmount);

        // --- Flat / pivot buffer (Tab 4) ---
        // FIX 5: the flat buffer had no primary-key value and relied on
        // Insert(true). AutoIncrement does not work on temporary records,
        // so this now assigns an explicit Entry No.
        ReconFlatBuffer.Init();
        ReconFlatBuffer."Entry No." := GetNextFlatBufEntryNo();
        ReconFlatBuffer."Period Start Date" := PeriodStart;
        ReconFlatBuffer."Period Label" := PeriodLabel;
        ReconFlatBuffer."Account No." := GLAccount."No.";
        ReconFlatBuffer."Account Name" := GLAccount.Name;
        ReconFlatBuffer."Account Category" := AccCategory;
        ReconFlatBuffer.Source := SourceText;
        ReconFlatBuffer."Party Type" := PartyType;
        ReconFlatBuffer."Party No." := PartyNo;
        ReconFlatBuffer."Party Name" := PartyName;
        ReconFlatBuffer."Is Subledger" := IsSubledger;
        ReconFlatBuffer."G/L Entry No." := GLEntry."Entry No.";
        ReconFlatBuffer."Document No." := GLEntry."Document No.";
        ReconFlatBuffer."Posting Date" := GLEntry."Posting Date";
        ReconFlatBuffer.Amount := EntryAmount;
        ReconFlatBuffer.Insert(false);
    end;

    local procedure NormalizeAmount(AccCategory: Enum "Recon Account Category"; Amount: Decimal): Decimal
    begin
        if AccCategory = AccCategory::Revenue then
            exit(-Amount);
        exit(Amount);
    end;

    local procedure PassesStrictCheck(GLEntry: Record "G/L Entry"): Boolean
    begin
        if not StrictLineValidation then
            exit(true);
        exit(LineHasGLAccount(
            GLEntry."Document No.", GLEntry."Document Type", GLEntry."Source Type", GLEntry."G/L Account No."));
    end;

    local procedure PassesStrictCheckForType(GLEntry: Record "G/L Entry"; PartyType: Enum "Recon Party Type"): Boolean
    var
        SourceType: Enum "Gen. Journal Source Type";
    begin
        if not StrictLineValidation then
            exit(true);
        case PartyType of
            PartyType::Vendor:
                SourceType := SourceType::Vendor;
            PartyType::Customer:
                SourceType := SourceType::Customer;
            else
                exit(false);
        end;
        exit(LineHasGLAccount(GLEntry."Document No.", GLEntry."Document Type", SourceType, GLEntry."G/L Account No."));
    end;

    local procedure AccumulatePartyBuffer(PartyType: Enum "Recon Party Type"; PartyNo: Code[20]; PartyName: Text[100]; PeriodStart: Date; PeriodLabel: Text[10]; Amount: Decimal)
    begin
        ReconPartyBuffer.Reset();
        if ReconPartyBuffer.Get(PartyType, PartyNo, PeriodStart) then begin
            ReconPartyBuffer.Amount += Amount;
            ReconPartyBuffer.Modify(false);
        end else begin
            ReconPartyBuffer.Init();
            ReconPartyBuffer."Party Type" := PartyType;
            ReconPartyBuffer."Party No." := PartyNo;
            ReconPartyBuffer."Party Name" := PartyName;
            ReconPartyBuffer."Period Start Date" := PeriodStart;
            ReconPartyBuffer."Period Label" := PeriodLabel;
            ReconPartyBuffer.Amount := Amount;
            ReconPartyBuffer.Insert(false);
        end;
    end;

    // ------------------------------------------------------------------
    // DOCUMENT LOOKUPS -- the actual "tie to subledger" logic
    // ------------------------------------------------------------------
    local procedure TryFindVendorDocument(GLEntry: Record "G/L Entry"; var VendNo: Code[20]; var VendName: Text[100]; var SourceText: Text[50]): Boolean
    var
        PurchInvHeader: Record "Purch. Inv. Header";
        PurchCrMemoHdr: Record "Purch. Cr. Memo Hdr.";
    begin
        case GLEntry."Document Type" of
            GLEntry."Document Type"::Invoice:
                if PurchInvHeader.Get(GLEntry."Document No.") then begin
                    VendNo := PurchInvHeader."Buy-from Vendor No.";
                    VendName := PurchInvHeader."Buy-from Vendor Name";
                    SourceText := 'Vendor Invoice (Purchase)';
                    exit(true);
                end;
            GLEntry."Document Type"::"Credit Memo":
                if PurchCrMemoHdr.Get(GLEntry."Document No.") then begin
                    VendNo := PurchCrMemoHdr."Buy-from Vendor No.";
                    VendName := PurchCrMemoHdr."Buy-from Vendor Name";
                    SourceText := 'Vendor Credit Memo (Purchase)';
                    exit(true);
                end;
        end;
        exit(false);
    end;

    local procedure TryFindCustomerDocument(GLEntry: Record "G/L Entry"; var CustNo: Code[20]; var CustName: Text[100]; var SourceText: Text[50]): Boolean
    var
        SalesInvHeader: Record "Sales Invoice Header";
        SalesCrMemoHeader: Record "Sales Cr.Memo Header";
    begin
        case GLEntry."Document Type" of
            GLEntry."Document Type"::Invoice:
                if SalesInvHeader.Get(GLEntry."Document No.") then begin
                    CustNo := SalesInvHeader."Sell-to Customer No.";
                    CustName := SalesInvHeader."Sell-to Customer Name";
                    SourceText := 'Customer Invoice (Sales)';
                    exit(true);
                end;
            GLEntry."Document Type"::"Credit Memo":
                if SalesCrMemoHeader.Get(GLEntry."Document No.") then begin
                    CustNo := SalesCrMemoHeader."Sell-to Customer No.";
                    CustName := SalesCrMemoHeader."Sell-to Customer Name";
                    SourceText := 'Customer Credit Memo (Sales)';
                    exit(true);
                end;
        end;
        exit(false);
    end;

    local procedure LineHasGLAccount(DocNo: Code[20]; DocType: Enum "Gen. Journal Document Type"; SourceType: Enum "Gen. Journal Source Type"; AccountNo: Code[20]): Boolean
    var
        PurchInvLine: Record "Purch. Inv. Line";
        PurchCrMemoLine: Record "Purch. Cr. Memo Line";
        SalesInvLine: Record "Sales Invoice Line";
        SalesCrMemoLine: Record "Sales Cr.Memo Line";
    begin
        case SourceType of
            SourceType::Vendor:
                case DocType of
                    DocType::Invoice:
                        begin
                            PurchInvLine.SetRange("Document No.", DocNo);
                            PurchInvLine.SetRange(Type, PurchInvLine.Type::"G/L Account");
                            PurchInvLine.SetRange("No.", AccountNo);
                            exit(not PurchInvLine.IsEmpty());
                        end;
                    DocType::"Credit Memo":
                        begin
                            PurchCrMemoLine.SetRange("Document No.", DocNo);
                            PurchCrMemoLine.SetRange(Type, PurchCrMemoLine.Type::"G/L Account");
                            PurchCrMemoLine.SetRange("No.", AccountNo);
                            exit(not PurchCrMemoLine.IsEmpty());
                        end;
                end;
            SourceType::Customer:
                case DocType of
                    DocType::Invoice:
                        begin
                            SalesInvLine.SetRange("Document No.", DocNo);
                            SalesInvLine.SetRange(Type, SalesInvLine.Type::"G/L Account");
                            SalesInvLine.SetRange("No.", AccountNo);
                            exit(not SalesInvLine.IsEmpty());
                        end;
                    DocType::"Credit Memo":
                        begin
                            SalesCrMemoLine.SetRange("Document No.", DocNo);
                            SalesCrMemoLine.SetRange(Type, SalesCrMemoLine.Type::"G/L Account");
                            SalesCrMemoLine.SetRange("No.", AccountNo);
                            exit(not SalesCrMemoLine.IsEmpty());
                        end;
                end;
        end;
        exit(false);
    end;

    // ------------------------------------------------------------------
    // EXCEL EXPORT
    // ------------------------------------------------------------------
    local procedure ExportToExcel(MonthLabel: array[6] of Text[10])
    begin
        // DIAGNOSTIC WRAPPER: the platform's own error for any failure in
        // this pipeline is the generic "Could not create the Excel
        // workbook." dialog, which hides the actual cause. Running the
        // whole export inside a TryFunction lets us surface the real
        // .NET/OpenXML exception text via GetLastErrorText() so the next
        // failure tells us exactly what to fix (bad characters in a text
        // field, row-count limit, sheet-name clash, etc.).
        if not TryExportToExcel(MonthLabel) then
            Error('Excel export failed at step "%1": %2', LastStepTxt, GetLastErrorText());
    end;

    // Global (not local) so its value survives the rollback that
    // TryFunction performs on failure -- only DATABASE changes are
    // rolled back, plain variables are not.
    var
        LastStepTxt: Text;

    [TryFunction]
    local procedure TryExportToExcel(MonthLabel: array[6] of Text[10])
    var
        TempExcelBuffer: Record "Excel Buffer" temporary;
    begin
        LastStepTxt := 'CreateNewBook';
        TempExcelBuffer.CreateNewBook('GL Account Balance');

        LastStepTxt := 'WriteGLBalanceSheet';
        WriteGLBalanceSheet(TempExcelBuffer, MonthLabel);
        LastStepTxt := 'WriteSheet (GL Account Balance)';
        TempExcelBuffer.WriteSheet('GL Account Balance', CompanyName(), UserId());
        TempExcelBuffer.DeleteAll();
        tempExcelBuffer.ClearNewRow();

        LastStepTxt := 'StartNewSheet (Subledger Reconciliation)';
        StartNewSheet(TempExcelBuffer, 'Subledger Reconciliation');
        LastStepTxt := 'WriteReconSheet';
        WriteReconSheet(TempExcelBuffer);
        LastStepTxt := 'WriteSheet (Subledger Reconciliation)';
        TempExcelBuffer.WriteSheet('Subledger Reconciliation', CompanyName(), UserId());
        TempExcelBuffer.DeleteAll();
        tempExcelBuffer.ClearNewRow();

        LastStepTxt := 'StartNewSheet (Vendor-Customer Comparison)';
        StartNewSheet(TempExcelBuffer, 'Vendor-Customer Comparison');
        LastStepTxt := 'WritePartySheet';
        WritePartySheet(TempExcelBuffer, MonthLabel);
        LastStepTxt := 'WriteSheet (Vendor-Customer Comparison)';
        TempExcelBuffer.WriteSheet('Vendor-Customer Comparison', CompanyName(), UserId());
        TempExcelBuffer.DeleteAll();
        tempExcelBuffer.ClearNewRow();

        LastStepTxt := 'StartNewSheet (Flat Data)';
        StartNewSheet(TempExcelBuffer, 'Flat Data (Pivot Source)');
        LastStepTxt := 'WriteFlatSheet';
        WriteFlatSheet(TempExcelBuffer);
        LastStepTxt := 'WriteSheet (Flat Data)';
        TempExcelBuffer.WriteSheet('Flat Data (Pivot Source)', CompanyName(), UserId());
        TempExcelBuffer.DeleteAll();
        tempExcelBuffer.ClearNewRow();

        LastStepTxt := 'CloseBook';
        TempExcelBuffer.SetFriendlyFilename('GL Subledger Reconciliation');
        TempExcelBuffer.CloseBook();

        LastStepTxt := 'OpenExcel';
        TempExcelBuffer.OpenExcel();
    end;

    // FIX 6: the original version wrote all four sheets on top of each
    // other. Clearing the rows is not enough -- you must also add/select
    // the next sheet in the workbook AND reset the buffer's internal row
    // counter, otherwise sheet 2 starts hundreds of blank rows down.
    // This is the only version-sensitive call in the codeunit: if
    // ClearNewRow() is not exposed in your BC version, remove that one
    // line only.
    local procedure StartNewSheet(var TempExcelBuffer: Record "Excel Buffer" temporary; SheetName: Text)
    begin
        TempExcelBuffer.SelectOrAddSheet(SheetName);
        // TempExcelBuffer.Reset();

    end;

    local procedure NewRow(var TempExcelBuffer: Record "Excel Buffer" temporary)
    begin
        TempExcelBuffer.NewRow();
    end;

    // FIX 7: AddColumn's 3rd parameter is CommentText and the 7th is
    // NumberFormat. The original passed '#,##0.00;(#,##0.00)' into the
    // comment slot, so every amount came out unformatted (and carried a
    // cell comment). Signature used here:
    //   AddColumn(Value, IsFormula, Comment, Bold, Italic, Underline,
    //             NumberFormat, CellType)
    local procedure AddTextCell(var TempExcelBuffer: Record "Excel Buffer" temporary; Value: Text; Bold: Boolean)
    begin
        TempExcelBuffer.AddColumn(Value, false, '', Bold, false, false, '', TempExcelBuffer."Cell Type"::Text);
    end;

    local procedure AddNumberCell(var TempExcelBuffer: Record "Excel Buffer" temporary; Value: Decimal; Bold: Boolean)
    begin
        TempExcelBuffer.AddColumn(
            Value, false, '', Bold, false, false, '#,##0.00;(#,##0.00)', TempExcelBuffer."Cell Type"::Number);
    end;

    local procedure AddDateCell(var TempExcelBuffer: Record "Excel Buffer" temporary; Value: Date)
    begin
        TempExcelBuffer.AddColumn(Value, false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Date);
    end;
    // MoM % cells. Value is stored as a FRACTION (0.032) with a percent
    // number format, so Excel shows 3.2% and the cell stays numeric for
    // conditional formatting / further formulas.
    local procedure AddPercentCell(var TempExcelBuffer: Record "Excel Buffer" temporary; Value: Decimal; Bold: Boolean)
    begin
        TempExcelBuffer.AddColumn(
            Value, false, '', Bold, false, false, '0.0%;(0.0%)', TempExcelBuffer."Cell Type"::Number);
    end;

    local procedure AddBlankCell(var TempExcelBuffer: Record "Excel Buffer" temporary)
    begin
        TempExcelBuffer.AddColumn('', false, '', false, false, false, '', TempExcelBuffer."Cell Type"::Text);
    end;

    // Abs() on the denominator so a contra account (e.g. Sales Returns,
    // which is negative in both months) still reads the right direction:
    // -2,103 -> -2,315 is a (10.1%) move, not +10.1%.
    local procedure TryMoM(Curr: Decimal; Prev: Decimal; var Pct: Decimal): Boolean
    begin
        Clear(Pct);
        if Prev = 0 then
            exit(false);
        Pct := (Curr - Prev) / Abs(Prev);
        exit(true);
    end;

    local procedure MoMHeader(MonthLabel: Text[10]): Text
    begin
        // "Mar-26" -> "MoM % Mar"
        exit('MoM % ' + CopyStr(MonthLabel, 1, 3));
    end;

    local procedure WriteMoMCells(var TempExcelBuffer: Record "Excel Buffer" temporary; MonthAmt: array[6] of Decimal)
    var
        i: Integer;
        Pct: Decimal;
    begin
        for i := 2 to 6 do
            if TryMoM(MonthAmt[i], MonthAmt[i - 1], Pct) then
                AddPercentCell(TempExcelBuffer, Pct, false)
            else
                AddBlankCell(TempExcelBuffer);   // no prior-month base
    end;
    // ------------------------------------------------------------------
    // SHEET 1: G/L Account Balance -- one row per account, six columns
    // ------------------------------------------------------------------
    // FIX 8: the original looped the buffer with FindSet and then called
    // SetRange/FindFirst on the SAME record variable inside the loop.
    // That destroyed the loop position (and CalcSums with no account
    // filter summed every account), so the sheet produced six near-
    // identical rows per account with wrong totals. Both grouped sheets
    // are now written in a single ordered pass -- no nested navigation.
    local procedure WriteGLBalanceSheet(var TempExcelBuffer: Record "Excel Buffer" temporary; MonthLabel: array[6] of Text[10])
    var
        MonthAmt: array[6] of Decimal;
        CurrAccNo: Code[20];
        CurrAccName: Text[100];
        CurrCategory: Text[30];
        i: Integer;
        RowOpen: Boolean;
    begin
        NewRow(TempExcelBuffer);
        AddTextCell(TempExcelBuffer, 'Account No.', true);
        AddTextCell(TempExcelBuffer, 'Account Name', true);
        AddTextCell(TempExcelBuffer, 'Category', true);

        for i := 1 to 6 do
            AddTextCell(TempExcelBuffer, MonthLabel[i], true);
        AddTextCell(TempExcelBuffer, '6-Mo Total', true);
        AddTextCell(TempExcelBuffer, 'Avg / Month', true);     // ADD THIS
        for i := 2 to 6 do                                     // ADD THIS
            AddTextCell(TempExcelBuffer, MoMHeader(MonthLabel[i]), true);



        ReconGLBuffer.Reset();
        ReconGLBuffer.SetCurrentKey("Account No.", "Period Start Date");
        if ReconGLBuffer.FindSet() then
            repeat
                if RowOpen and (ReconGLBuffer."Account No." <> CurrAccNo) then begin
                    WriteBalanceRow(TempExcelBuffer, CurrAccNo, CurrAccName, CurrCategory, MonthAmt);
                    Clear(MonthAmt);
                end;
                CurrAccNo := ReconGLBuffer."Account No.";
                CurrAccName := ReconGLBuffer."Account Name";
                CurrCategory := Format(ReconGLBuffer."Account Category");
                RowOpen := true;

                i := MonthIndex(MonthLabel, ReconGLBuffer."Period Label");
                if i > 0 then
                    MonthAmt[i] += ReconGLBuffer."GL Amount";
            until ReconGLBuffer.Next() = 0;

        if RowOpen then
            WriteBalanceRow(TempExcelBuffer, CurrAccNo, CurrAccName, CurrCategory, MonthAmt);
    end;


    local procedure WriteBalanceRow(var TempExcelBuffer: Record "Excel Buffer" temporary; AccNo: Code[20]; AccName: Text[100]; Category: Text[30]; MonthAmt: array[6] of Decimal)
    var
        i: Integer;
        Total: Decimal;
    begin
        NewRow(TempExcelBuffer);
        AddTextCell(TempExcelBuffer, AccNo, false);
        AddTextCell(TempExcelBuffer, AccName, false);
        AddTextCell(TempExcelBuffer, Category, false);
        for i := 1 to 6 do begin
            AddNumberCell(TempExcelBuffer, MonthAmt[i], false);
            Total += MonthAmt[i];
        end;
        AddNumberCell(TempExcelBuffer, Total, true);
        AddNumberCell(TempExcelBuffer, Total / 6, false);
        WriteMoMCells(TempExcelBuffer, MonthAmt);
    end;
    // ------------------------------------------------------------------
    // SHEET 2: Subledger Reconciliation -- Account x Month detail
    // ------------------------------------------------------------------
    local procedure WriteReconSheet(var TempExcelBuffer: Record "Excel Buffer" temporary)
    begin
        NewRow(TempExcelBuffer);
        AddTextCell(TempExcelBuffer, 'Account No.', true);
        AddTextCell(TempExcelBuffer, 'Account Name', true);
        AddTextCell(TempExcelBuffer, 'Category', true);
        AddTextCell(TempExcelBuffer, 'Period', true);
        AddTextCell(TempExcelBuffer, 'GL Amount', true);
        AddTextCell(TempExcelBuffer, 'Subledger Amount', true);
        AddTextCell(TempExcelBuffer, 'Without Vendor', true);
        AddTextCell(TempExcelBuffer, 'Without Customer', true);
        AddTextCell(TempExcelBuffer, 'Variance (should be 0)', true);

        ReconGLBuffer.Reset();
        ReconGLBuffer.SetCurrentKey("Account No.", "Period Start Date");
        if ReconGLBuffer.FindSet() then
            repeat
                NewRow(TempExcelBuffer);
                AddTextCell(TempExcelBuffer, ReconGLBuffer."Account No.", false);
                AddTextCell(TempExcelBuffer, ReconGLBuffer."Account Name", false);
                AddTextCell(TempExcelBuffer, Format(ReconGLBuffer."Account Category"), false);
                AddTextCell(TempExcelBuffer, ReconGLBuffer."Period Label", false);
                AddNumberCell(TempExcelBuffer, ReconGLBuffer."GL Amount", false);
                AddNumberCell(TempExcelBuffer, ReconGLBuffer."Subledger Amount", false);
                AddNumberCell(TempExcelBuffer, ReconGLBuffer."Without Vendor Amount", false);
                AddNumberCell(TempExcelBuffer, ReconGLBuffer."Without Customer Amount", false);
                AddNumberCell(TempExcelBuffer, ReconGLBuffer.Variance, false);
            until ReconGLBuffer.Next() = 0;
    end;

    // ------------------------------------------------------------------
    // SHEET 3: Vendor / Customer comparison -- one row per party
    // ------------------------------------------------------------------
    local procedure WritePartySheet(var TempExcelBuffer: Record "Excel Buffer" temporary; MonthLabel: array[6] of Text[10])
    var
        MonthAmt: array[6] of Decimal;
        CurrType: Enum "Recon Party Type";
        CurrNo: Code[20];
        CurrName: Text[100];
        i: Integer;
        RowOpen: Boolean;
    begin
        NewRow(TempExcelBuffer);
        AddTextCell(TempExcelBuffer, 'Party Type', true);
        AddTextCell(TempExcelBuffer, 'No.', true);
        AddTextCell(TempExcelBuffer, 'Name', true);
        for i := 1 to 6 do
            AddTextCell(TempExcelBuffer, MonthLabel[i], true);
        AddTextCell(TempExcelBuffer, '6-Mo Total', true);
        AddTextCell(TempExcelBuffer, 'Avg / Month', true);
        for i := 2 to 6 do
            AddTextCell(TempExcelBuffer, MoMHeader(MonthLabel[i]), true);
        ReconPartyBuffer.Reset();
        ReconPartyBuffer.SetCurrentKey("Party Type", "Party No.", "Period Start Date");
        if ReconPartyBuffer.FindSet() then
            repeat
                if RowOpen and
                   ((ReconPartyBuffer."Party Type" <> CurrType) or (ReconPartyBuffer."Party No." <> CurrNo))
                then begin
                    WritePartyRow(TempExcelBuffer, CurrType, CurrNo, CurrName, MonthAmt);
                    Clear(MonthAmt);
                end;
                CurrType := ReconPartyBuffer."Party Type";
                CurrNo := ReconPartyBuffer."Party No.";
                CurrName := ReconPartyBuffer."Party Name";
                RowOpen := true;

                i := MonthIndex(MonthLabel, ReconPartyBuffer."Period Label");
                if i > 0 then
                    MonthAmt[i] += ReconPartyBuffer.Amount;
            until ReconPartyBuffer.Next() = 0;

        if RowOpen then
            WritePartyRow(TempExcelBuffer, CurrType, CurrNo, CurrName, MonthAmt);
    end;

    local procedure WritePartyRow(var TempExcelBuffer: Record "Excel Buffer" temporary; PartyType: Enum "Recon Party Type"; PartyNo: Code[20]; PartyName: Text[100]; MonthAmt: array[6] of Decimal)
    var
        i: Integer;
        Total: Decimal;
    begin
        NewRow(TempExcelBuffer);
        AddTextCell(TempExcelBuffer, Format(PartyType), false);
        AddTextCell(TempExcelBuffer, PartyNo, false);
        AddTextCell(TempExcelBuffer, PartyName, false);
        for i := 1 to 6 do begin
            AddNumberCell(TempExcelBuffer, MonthAmt[i], false);
            Total += MonthAmt[i];
        end;
        AddNumberCell(TempExcelBuffer, Total, true);
        WriteMoMCells(TempExcelBuffer, MonthAmt);
    end;

    // ------------------------------------------------------------------
    // SHEET 4: Flat data (pivot source)
    // ------------------------------------------------------------------
    local procedure WriteFlatSheet(var TempExcelBuffer: Record "Excel Buffer" temporary)
    begin
        NewRow(TempExcelBuffer);
        AddTextCell(TempExcelBuffer, 'Period', true);
        AddTextCell(TempExcelBuffer, 'Account No.', true);
        AddTextCell(TempExcelBuffer, 'Account Name', true);
        AddTextCell(TempExcelBuffer, 'Account Category', true);
        AddTextCell(TempExcelBuffer, 'Source', true);
        AddTextCell(TempExcelBuffer, 'Party Type', true);
        AddTextCell(TempExcelBuffer, 'Party No.', true);
        AddTextCell(TempExcelBuffer, 'Party Name', true);
        AddTextCell(TempExcelBuffer, 'Is Subledger', true);
        AddTextCell(TempExcelBuffer, 'Document No.', true);
        AddTextCell(TempExcelBuffer, 'G/L Entry No.', true);
        AddTextCell(TempExcelBuffer, 'Posting Date', true);
        AddTextCell(TempExcelBuffer, 'Amount', true);

        ReconFlatBuffer.Reset();
        ReconFlatBuffer.SetCurrentKey("Period Start Date", "Account No.");
        if ReconFlatBuffer.FindSet() then
            repeat
                NewRow(TempExcelBuffer);
                AddTextCell(TempExcelBuffer, ReconFlatBuffer."Period Label", false);
                AddTextCell(TempExcelBuffer, ReconFlatBuffer."Account No.", false);
                AddTextCell(TempExcelBuffer, ReconFlatBuffer."Account Name", false);
                AddTextCell(TempExcelBuffer, Format(ReconFlatBuffer."Account Category"), false);
                AddTextCell(TempExcelBuffer, ReconFlatBuffer.Source, false);
                AddTextCell(TempExcelBuffer, Format(ReconFlatBuffer."Party Type"), false);
                AddTextCell(TempExcelBuffer, ReconFlatBuffer."Party No.", false);
                AddTextCell(TempExcelBuffer, ReconFlatBuffer."Party Name", false);
                AddTextCell(TempExcelBuffer, Format(ReconFlatBuffer."Is Subledger"), false);
                AddTextCell(TempExcelBuffer, ReconFlatBuffer."Document No.", false);
                AddTextCell(TempExcelBuffer, Format(ReconFlatBuffer."G/L Entry No."), false);
                AddDateCell(TempExcelBuffer, ReconFlatBuffer."Posting Date");
                AddNumberCell(TempExcelBuffer, ReconFlatBuffer.Amount, false);
            until ReconFlatBuffer.Next() = 0;
    end;
}
