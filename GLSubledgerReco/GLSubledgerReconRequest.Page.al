page 50100 "GL Subledger Recon Request"
{
    Caption = 'GL to Subledger Reconciliation';
    PageType = StandardDialog;
    ApplicationArea = All;
    UsageCategory = ReportsAndAnalysis;

    layout
    {
        area(Content)
        {
            group(Options)
            {
                Caption = 'Options';

                field(AsOfDate; AsOfDate)
                {
                    ApplicationArea = All;
                    Caption = 'As of Date';
                    ToolTip = 'The report covers the month containing this date plus the five preceding calendar months.';
                }
                field(StrictLineValidation; StrictLineValidation)
                {
                    ApplicationArea = All;
                    Caption = 'Strict Line Validation';
                    ToolTip = 'Counts a posting as subledger-backed only when the posted document carries a G/L Account type line for that exact account. Leave blank unless you specifically want to isolate G/L-Account-type invoice lines.';
                }
            }
        }
    }

    trigger OnOpenPage()
    begin
        if AsOfDate = 0D then
            AsOfDate := WorkDate();
    end;

    trigger OnQueryClosePage(CloseAction: Action): Boolean
    var
        ReconMgt: Codeunit "GL Subledger Recon Mgt.";
    begin
        if CloseAction = Action::OK then
            ReconMgt.RunExport(AsOfDate, StrictLineValidation);
        exit(true);
    end;

    var
        AsOfDate: Date;
        StrictLineValidation: Boolean;
}
