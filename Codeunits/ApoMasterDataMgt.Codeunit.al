codeunit 50002 "Apo Master Data Mgt."
{
    trigger OnRun()
    begin
    end;

    var
        VendorErrorMessageLogTmp_GT: Record "Error Message" temporary;

    procedure SetVendorStatusDispatcher(var Vendor_PT: Record Vendor; ShowError_P: Boolean; Status_P: Option Open,Check,Released)
    begin
        VendorErrorMessageLogTmp_GT.Reset();
        VendorErrorMessageLogTmp_GT.DeleteAll();
        case Status_P of
            Status_P::Open:
                SetVendorStatusOpen(Vendor_PT);
            Status_P::Check:
                begin
                    SetVendorStatusCheck(Vendor_PT, ShowError_P, false);
                    SetVendorStatusReleased(Vendor_PT, false, true);
                end;
            Status_P::Released:
                SetVendorStatusReleased(Vendor_PT, ShowError_P, false);
            else begin
                Vendor_PT.Status := Status_P;
                Vendor_PT.FieldError(Status);
            end;
        end;
    end;

    local procedure SetVendorStatusOpen(var Vendor_PT: Record Vendor)
    begin
        if Vendor_PT.Status = Vendor_PT.Status::Open then
            exit;
        Vendor_PT.Status := Vendor_PT.Status::Open;
        Vendor_PT.Modify();
    end;

    local procedure SetVendorStatusCheck(var Vendor_PT: Record Vendor; ShowError_P: Boolean; CalledFromRelease_P: Boolean): Boolean
    var
        GeneralLedgerSetup_LT: Record "General Ledger Setup";
        DefaultDimension_LT: Record "Default Dimension";
        ConfigTemplateHeader_LT: Record "Config. Template Header";
        DateFormula_L: DateFormula;
    begin
        if Vendor_PT."Created with Cfg. Tmpl. Hdr." <> '' then begin
            if ConfigTemplateHeader_LT.Get(Vendor_PT."Created with Cfg. Tmpl. Hdr.") then begin
                if ConfigTemplateHeader_LT.Enabled then begin
                    SetVendorStatusCheckWithTemplateCfgHdr(Vendor_PT, ConfigTemplateHeader_LT, 2);
                    if VendorErrorMessageLogTmp_GT.HasErrors(false) then begin
                        if ShowError_P then begin
                            if not CalledFromRelease_P then
                                VendorErrorMessageLogTmp_GT.ShowErrorMessages(false);
                        end;
                        exit(false);
                    end else begin
                        if not CalledFromRelease_P then begin
                            SetVendorStatusCheckWithTemplateCfgHdr(Vendor_PT, ConfigTemplateHeader_LT, 3);
                            if not VendorErrorMessageLogTmp_GT.HasErrors(false) then
                                Vendor_PT.Status := Vendor_PT.Status::Released
                            else
                                Vendor_PT.Status := Vendor_PT.Status::Check;
                            Vendor_PT.Modify();
                        end;
                        exit(true);
                    end;
                end;
            end;
        end;
        if ShowError_P then begin
            Vendor_PT.TestField("Location Code");
            Vendor_PT.TestField("Vendor Type");
            Vendor_PT.TestField("Purchaser Code");
            Vendor_PT.TestField("Country/Region Code");
            Vendor_PT.TestField("Our Account No.");
            Vendor_PT.TestField("Document Sending Profile");
            Vendor_PT.TestField("E-Mail");
            Vendor_PT.TestField("Lead Time Calculation");
            if Vendor_PT."Price and Discount Source" = Vendor_PT."Price and Discount Source"::"Buy-from Vendor No." then
                Vendor_PT.TestField("Pay-to Vendor No.");
            DefaultDimension_LT.Get(23, Vendor_PT."No.", GeneralLedgerSetup_LT."Shortcut Dimension 3 Code");
        end else begin
            if Vendor_PT."Purchaser Code" = '' then
                exit(false);
            if Vendor_PT."Country/Region Code" = '' then
                exit(false);
            if Vendor_PT."Our Account No." = '' then
                exit(false);
            if Vendor_PT."Document Sending Profile" = '' then
                exit(false);
            if Vendor_PT."E-Mail" = '' then
                exit(false);
        end;

        if not CalledFromRelease_P then begin
            Vendor_PT.Status := Vendor_PT.Status::Check;
            Vendor_PT.Modify();
        end;

        exit(true);
    end;

    local procedure SetVendorStatusReleased(var Vendor_PT: Record Vendor; ShowError_P: Boolean; CalledFromCheck_P: Boolean): Boolean
    var
        ConfigTemplateHeader_LT: Record "Config. Template Header";
        ApoFunctionsMgt_LC: Codeunit "Apo Functions Mgt.";
    begin
        if not CalledFromCheck_P then begin
            Clear(ApoFunctionsMgt_LC);
            ApoFunctionsMgt_LC.CheckFiCoUser(true);
        end;
        if Vendor_PT."Created with Cfg. Tmpl. Hdr." <> '' then begin
            if ConfigTemplateHeader_LT.Get(Vendor_PT."Created with Cfg. Tmpl. Hdr.") then begin
                if ConfigTemplateHeader_LT.Enabled then begin
                    SetVendorStatusCheckWithTemplateCfgHdr(Vendor_PT, ConfigTemplateHeader_LT, 2);
                    SetVendorStatusCheckWithTemplateCfgHdr(Vendor_PT, ConfigTemplateHeader_LT, 3);
                    if VendorErrorMessageLogTmp_GT.HasErrors(false) then begin
                        if ShowError_P then begin
                            if not CalledFromCheck_P then
                                VendorErrorMessageLogTmp_GT.ShowErrorMessages(false);
                        end;
                        exit(false);
                    end else begin
                        if not CalledFromCheck_P then begin
                            Vendor_PT.Status := Vendor_PT.Status::Released;
                            Vendor_PT.Modify();
                        end;
                        exit(true);
                    end;
                end;
            end;
        end;
        SetVendorStatusCheck(Vendor_PT, ShowError_P, true);
        if ShowError_P then begin
            Vendor_PT.TestField("Gen. Bus. Posting Group");
            Vendor_PT.TestField("VAT Bus. Posting Group");
            Vendor_PT.TestField("Vendor Posting Group");
            Vendor_PT.TestField("VAT Registration No.");
            Vendor_PT.TestField("Payment Terms Code");
            Vendor_PT.TestField("Payment Method Code");
        end else begin
            if Vendor_PT."Gen. Bus. Posting Group" = '' then
                exit(false);
            if Vendor_PT."VAT Bus. Posting Group" = '' then
                exit(false);
            if Vendor_PT."Vendor Posting Group" = '' then
                exit(false);
            if Vendor_PT."VAT Registration No." = '' then
                exit(false);
            if Vendor_PT."Payment Terms Code" = '' then
                exit(false);
            if Vendor_PT."Payment Method Code" = '' then
                exit(false);
        end;

        Vendor_PT.Status := Vendor_PT.Status::Released;
        Vendor_PT.Modify();

        exit(true);
    end;

    local procedure SetVendorStatusCheckWithTemplateCfgHdr(var Vendor_PT: Record Vendor; ConfigTemplateHeader_PT: Record "Config. Template Header"; NewStatus_P: Option Open,Check,Released)
    var
        ConfigTemplateLine_LT: Record "Config. Template Line";
        DataTypeMgt_LC: Codeunit "Data Type Management";
        ApoFunctionsMgt_LC: Codeunit "Apo Functions Mgt.";
        RecRef_L: RecordRef;
        RecRefChk_L: RecordRef;
        FldRef_L: FieldRef;
        CurrentFldNo_L: Integer;
        i: Integer;
        FieldIsEmptyErrTxt: Label 'Field %1 is empty';
        FieldIsEmptyWithConditErrTxt: Label 'Field %1 is empty (Condition: %2)';
    begin
        DataTypeMgt_LC.GetRecordRef(Vendor_PT, RecRef_L);
        for i := 1 to RecRef_L.FieldCount() do begin
            Clear(FldRef_L);
            FldRef_L := RecRef_L.FieldIndex(i);
            CurrentFldNo_L := FldRef_L.Number();
            if CheckFieldForMandatory(ConfigTemplateHeader_PT.Code, Database::Vendor, CurrentFldNo_L, NewStatus_P, ConfigTemplateLine_LT) then begin
                Clear(RecRefChk_L);
                if Format(ConfigTemplateLine_LT."Condit. Mandatory on Record") <> '' then begin
                    RecRefChk_L.Get(RecRef_L.RecordId());
                    RecRefChk_L.SetRecFilter();
                    ApoFunctionsMgt_LC.FilterRecRef(RecRefChk_L, Format(ConfigTemplateLine_LT."Condit. Mandatory on Record"), 0);
                    if RecRefChk_L.FindSet() then
                        if Format(FldRef_L.Value()) = '' then
                            VendorErrorMessageLogTmp_GT.LogDetailedMessage(Vendor_PT, FldRef_L.Number(), VendorErrorMessageLogTmp_GT."Message Type"::Error,
                                                                                                        CopyStr(StrSubstNo(FieldIsEmptyWithConditErrTxt, FldRef_L.Caption(), Format(ConfigTemplateLine_LT."Condit. Mandatory on Record")), 1, MaxStrLen(VendorErrorMessageLogTmp_GT.Message)), '', '');
                end else begin
                    if Format(FldRef_L.Value()) = '' then
                        VendorErrorMessageLogTmp_GT.LogDetailedMessage(Vendor_PT, FldRef_L.Number(), VendorErrorMessageLogTmp_GT."Message Type"::Error, StrSubstNo(FieldIsEmptyErrTxt, FldRef_L.Caption()), '', '');
                end;
            end;
        end;
    end;

    procedure CheckFieldForMandatory(DataTemplateCode_P: Code[10]; TableID_P: Integer; FieldID_P: Integer; MandatoryOnCreatedRecord_P: Option; var ConfigTemplateLine_PT: Record "Config. Template Line") FieldIsMandatory_L: Boolean
    begin
        Clear(ConfigTemplateLine_PT);
        ConfigTemplateLine_PT.SetRange("Data Template Code", DataTemplateCode_P);
        ConfigTemplateLine_PT.SetRange(Type, ConfigTemplateLine_PT.Type::Field);
        ConfigTemplateLine_PT.SetRange("Table ID", TableID_P);
        ConfigTemplateLine_PT.SetRange("Field ID", FieldID_P);
        ConfigTemplateLine_PT.SetRange("Mandatory on created Record", MandatoryOnCreatedRecord_P);
        exit(ConfigTemplateLine_PT.FindFirst());
    end;
}

