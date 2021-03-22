page 72001 bc2XLOAuthSetup
{
    
    ApplicationArea = All;
    Caption = 'bc2XLOAuthSetup';
    PageType = List;
    SourceTable = "OAuth 2.0 Setup";
    UsageCategory = Administration;
    
    layout
    {
        area(content)
        {
            repeater(General)
            {
                field(Code; Rec.Code)
                {
                    ApplicationArea = All;
                }
                field(Description; Rec.Description)
                {
                    ApplicationArea = All;
                }
                field("Access Token Due DateTime"; Rec."Access Token Due DateTime")
                {
                    ApplicationArea = All;
                }
                field("Access Token URL Path"; Rec."Access Token URL Path")
                {
                    ApplicationArea = All;
                }
                field("Access Token"; Rec."Access Token")
                {
                    ApplicationArea = All;
                }
                field("Activity Log ID"; Rec."Activity Log ID")
                {
                    ApplicationArea = All;
                }
                field("Authorization Response Type"; Rec."Authorization Response Type")
                {
                    ApplicationArea = All;
                }
                field("Authorization URL Path"; Rec."Authorization URL Path")
                {
                    ApplicationArea = All;
                }
                field("Client ID"; Rec."Client ID")
                {
                    ApplicationArea = All;
                }
                field("Client Secret"; Rec."Client Secret")
                {
                    ApplicationArea = All;
                }
                field("Daily Count"; Rec."Daily Count")
                {
                    ApplicationArea = All;
                }
                field("Daily Limit"; Rec."Daily Limit")
                {
                    ApplicationArea = All;
                }
                field("Latest Datetime"; Rec."Latest Datetime")
                {
                    ApplicationArea = All;
                }
                field("Redirect URL"; Rec."Redirect URL")
                {
                    ApplicationArea = All;
                }
                field("Refresh Token URL Path"; Rec."Refresh Token URL Path")
                {
                    ApplicationArea = All;
                }
                field("Refresh Token"; Rec."Refresh Token")
                {
                    ApplicationArea = All;
                }
                field("Service URL"; Rec."Service URL")
                {
                    ApplicationArea = All;
                }
                field("Token DataScope"; Rec."Token DataScope")
                {
                    ApplicationArea = All;
                }
                field(Scope; Rec.Scope)
                {
                    ApplicationArea = All;
                }
                field(Status; Rec.Status)
                {
                    ApplicationArea = All;
                }
                field(SystemCreatedAt; Rec.SystemCreatedAt)
                {
                    ApplicationArea = All;
                }
                field(SystemCreatedBy; Rec.SystemCreatedBy)
                {
                    ApplicationArea = All;
                }
                field(SystemId; Rec.SystemId)
                {
                    ApplicationArea = All;
                }
                field(SystemModifiedAt; Rec.SystemModifiedAt)
                {
                    ApplicationArea = All;
                }
                field(SystemModifiedBy; Rec.SystemModifiedBy)
                {
                    ApplicationArea = All;
                }
            }
        }
    }
    
}
