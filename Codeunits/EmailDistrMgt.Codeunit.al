codeunit 50059 "Email Distr. Mgt."
{
    // version ApoNL14.04
    Permissions = tabledata "Mailing Group" = rimd,
                  tabledata "Contact Mailing Group" = rimd;

    trigger OnRun()
    begin
    end;

    var
        GTS: array[10] of Text;
        DistributionCounter: Integer;
        SilentDistribution: Boolean;
        SaveFileDialogFilterMsg: Label 'PDF Files (*.pdf)|*.pdf';
        SaveFileDialogTitleMsg: Label 'Save PDF file';
        TextAttachmentFileName: Label '%1 - %2 %3.pdf', Comment = '%1 = Company Name. %2 = Document Type %3 = Invoice No.';
        TextRegExInfiniteLoop: Label 'A too big variable substitution loop in the text was found.';
        TextSetupNotFound: Label 'Could not find a setup for record %1 and report no. %2.';
        TextTypeNotSupported: Label 'The %1 for %2 is not supported.';

    procedure DistributeDocumentMails(FilteredHeaderDoc: Variant; ReportId: Integer; DistributionCode: Code[10])
    var
        DataTypeMgt: Codeunit "Data Type Management";
        RecRef: RecordRef;
    begin
        if ReportId = 0 then begin
            SearchReportSelection(FilteredHeaderDoc, ReportId, DistributionCode);
        end;
        if DistributionCode = '' then begin
            Error(TextSetupNotFound, Format(FilteredHeaderDoc), ReportId);
        end;
        DataTypeMgt.GetRecordRef(FilteredHeaderDoc, RecRef);
        if not RecRef.IsTemporary() then begin
            case RecRef.Count() of
                0:
                    exit;
                1:
                    DoDistributeDocumentMail(RecRef, ReportId, DistributionCode);
                else begin
                    DoDistributeBatchDocumentMail(RecRef, ReportId, DistributionCode);
                end;
            end;
        end else begin
            DoDistributeDocumentMail(RecRef, ReportId, DistributionCode);
        end;
    end;

    procedure DistributeDocumentMail(FilteredHeaderDoc: Variant; ReportId: Integer; DistributionCode: Code[10])
    var
        DataTypeMgt: Codeunit "Data Type Management";
        RecRef: RecordRef;
    begin
        DataTypeMgt.GetRecordRef(FilteredHeaderDoc, RecRef);
        RecRef.SetRecFilter();
        FilteredHeaderDoc := RecRef;
        DistributeDocumentMails(FilteredHeaderDoc, ReportId, DistributionCode);
    end;

    procedure DistributeDocumentMailsByReportSelection(FilteredHeaderDoc: Variant; ReportSelection: Record "Report Selections")
    begin
        DistributeDocumentMails(FilteredHeaderDoc, ReportSelection."Report ID", ReportSelection."Email Distribution Code");
    end;

    procedure DistributeDocumentMailsByCustReportSelection(FilteredHeaderDoc: Variant; CustReportSelection: Record "Custom Report Selection")
    begin
        DistributeDocumentMails(FilteredHeaderDoc, CustReportSelection."Report ID", CustReportSelection."Email Distribution Code");
    end;

    procedure DistributeRecordMail(EmailDistributionEntry_PT: Record "Email Distribution Entry"): Boolean
    var
        TempDistributionEntry: Record "Email Distribution Entry" temporary;
        EmailDistributionSetupLine: Record "Email Distribution Setup";
        DataTypeManagement_LC: Codeunit "Data Type Management";
        HeaderEntryMgt: Codeunit "Email Distr. Header Mgt.";
        RecRef_L: RecordRef;
        ExtraFilePathList_L: DotNet List_Of_T;
        ExtraFilePaths_LP: array[10] of Text;
        RequestPageParameters: Text;
        ReportID_L: Integer;
    begin
        if not RecRef_L.Get(EmailDistributionEntry_PT."Source Record ID") then
            exit(false);
        if not RecRef_L.IsTemporary() then begin
            TempDistributionEntry := EmailDistributionEntry_PT;
            TempDistributionEntry.Insert();
            EmailDistributionSetupLine.Get(TempDistributionEntry."Email Distribution Code", TempDistributionEntry."Language Code", TempDistributionEntry."Use for Type", TempDistributionEntry."Use for Code");
            // APO.002 SER 24.08.22 ...
            // APO.037 JUN 31.07.24 ...
            ExtraFilePathList_L := ExtraFilePathList_L.List;
            PrepareAddtionalAttachments(EmailDistributionSetupLine, TempDistributionEntry, RecRef_L, ExtraFilePathList_L);
            // statt:
            //PrepareAddtionalAttachments(EmailDistributionSetupLine,TempDistributionEntry,RecRef_L,ExtraFilePaths_L);
            // ... APO.037 JUN 31.07.24
            ReportID_L := 0;
            if TrySearchReportSelection(RecRef_L, ReportID_L, TempDistributionEntry."Email Distribution Code") then begin
                RecRef_L.SetRecFilter();
                // APO.037 JUN 31.07.24 ...
                exit(SendAsMail(RecRef_L, ReportID_L, EmailDistributionSetupLine, TempDistributionEntry, RequestPageParameters, ExtraFilePathList_L))
                // statt:
                //EXIT(SendAsMail(RecRef_L,ReportID_L,EmailDistributionSetupLine,TempDistributionEntry,RequestPageParameters,ExtraFilePaths_L))
                // ... APO.037 JUN 31.07.24
            end else
                // APO.037 JUN 31.07.24 ...
                exit(SendAsMail(RecRef_L, 0, EmailDistributionSetupLine, TempDistributionEntry, RequestPageParameters, ExtraFilePathList_L));
            // statt:
            //EXIT(SendAsMail(RecRef_L,0,EmailDistributionSetupLine,TempDistributionEntry,RequestPageParameters,ExtraFilePaths_L));
            // ... APO.037 JUN 31.07.24
            // ... APO.002 SER 24.08.22
        end;
    end;

    procedure PrepareDistrEntryAfterFinalizePostSalesPost(HeaderDoc_P: Variant; CustNo_P: Code[20]; ReportUsage_P: Integer) CreatedEntryNo_L: Integer
    var
        Customer_LT: Record Customer;
        DocumentSendingProfile_LT: Record "Document Sending Profile";
        ReportSelections_LT: Record "Report Selections";
        TempAttachReportSelections_LT: Record "Report Selections" temporary;
        TempEmailItem_LT: Record "Email Item" temporary;
        TempDistrEntry_LT: Record "Email Distribution Entry" temporary;
        EmailDistributionSetupLine_LT: Record "Email Distribution Setup";
        DataTypeManagement_LC: Codeunit "Data Type Management";
        EmailDistrHeaderMgt_LC: Codeunit "Email Distr. Header Mgt.";
        RecRef_L: RecordRef;
    begin
        Clear(RecRef_L);
        // APO.022 SER 10.08.23 ...
        if not DocumentSendingProfile_LT.Get(Customer_LT."Document Sending Profile") then
            exit;
        // DocumentSendingProfile_LT.GetDefaultForCustomer(CustNo_P,DocumentSendingProfile_LT);
        if DocumentSendingProfile_LT.LookupProfile(CustNo_P, false, false) then begin
            if DocumentSendingProfile_LT."E-Mail" = DocumentSendingProfile_LT."E-Mail"::No then
                exit;
            if not ReportSelections_LT.FindEmailAttachmentUsageForCust(ReportUsage_P, CustNo_P, TempAttachReportSelections_LT) then
                exit;
            if TempAttachReportSelections_LT."Email Distribution Code" = '' then
                exit;
            EmailDistrHeaderMgt_LC.GetHeaderEntry(HeaderDoc_P, TempDistrEntry_LT);
            DataTypeManagement_LC.GetRecordRef(HeaderDoc_P, RecRef_L);
            if TryGetDistributionLine(TempAttachReportSelections_LT."Email Distribution Code", TempDistrEntry_LT, EmailDistributionSetupLine_LT) then begin
                if TempDistrEntry_LT."Distribution Type" = TempDistrEntry_LT."Distribution Type"::Mail then begin
                    TempEmailItem_LT.ID := CreateGuid();
                    if EmailDistributionSetupLine_LT."Delay Send E-Mail" then begin
                        // APO.036 SER 10.07.24 ...
                        if (EmailDistributionSetupLine_LT."Delay Send E-Mail only if JQ") and not (CurrentClientType() in [ClientType::Background, ClientType::NAS]) then
                            exit(0);
                        PreCheckForFixedEMailRecipient(EmailDistributionSetupLine_LT, TempDistrEntry_LT); // APO.011 SER 12.09.22
                        if TryFillTempEmailItem(RecRef_L, 0, EmailDistributionSetupLine_LT, TempDistrEntry_LT, TempEmailItem_LT, '') then // only log, without report id, because with report the function will print it (at this moment not necessary)
                            exit(LogMailDistribution(TempEmailItem_LT, 0, GetPlannedSendingDateTime(EmailDistributionSetupLine_LT, 0DT)));
                        // ... APO.036 SER 10.07.24
                    end else
                        exit(0);
                end;
            end;
        end;
        // ... APO.010 SER 24.08.22
    end;

    local procedure "--- Do ---"()
    begin
    end;

    local procedure DoDistributeDocumentMail(RecRef: RecordRef; ReportId: Integer; DistributionCode: Code[10]): Boolean
    var
        TempDistributionEntry: Record "Email Distribution Entry" temporary;
        EmailDistributionSetupLine: Record "Email Distribution Setup";
        HeaderEntryMgt: Codeunit "Email Distr. Header Mgt.";
        ExtraFilePathList_L: DotNet List_Of_T;
        ExtraFilePaths_LP: array[10] of Text;
        RequestPageParameters: Text;
        DistributionType: Integer;
    begin
        HeaderEntryMgt.GetHeaderEntry(RecRef, TempDistributionEntry);
        GetDistributionLine(DistributionCode, TempDistributionEntry, EmailDistributionSetupLine);
        DistributionType := GetDistributionType(RecRef, EmailDistributionSetupLine);
        if EmailDistributionSetupLine."Show Request Page" and not SilentDistribution then begin
            RequestPageParameters := RunAndGetRequestPage(ReportId);
        end;
        case DistributionType of
            TempDistributionEntry."Distribution Type"::Mail:
                // APO.002 SER 24.08.22 ...
                begin
                    PreCheckForFixedEMailRecipient(EmailDistributionSetupLine, TempDistributionEntry); // APO.011 SER 12.09.22
                    // APO.037 JUN 31.07.24 ...
                    ExtraFilePathList_L := ExtraFilePathList_L.List;
                    PrepareAddtionalAttachments(EmailDistributionSetupLine, TempDistributionEntry, RecRef, ExtraFilePathList_L);
                    exit(SendAsMail(RecRef, ReportId, EmailDistributionSetupLine, TempDistributionEntry, RequestPageParameters, ExtraFilePathList_L));
                    // statt:
                    //PrepareAddtionalAttachments(EmailDistributionSetupLine,TempDistributionEntry,RecRef,ExtraFilePaths_L);
                    //EXIT(SendAsMail(RecRef,ReportId,EmailDistributionSetupLine,TempDistributionEntry,RequestPageParameters,ExtraFilePaths_L));
                    // ... APO.037 JUN 31.07.24
                end;
            // ... APO.002 SER 24.08.22
            TempDistributionEntry."Distribution Type"::PDF:
                exit(CreateAndDownloadPdf(RecRef, ReportId, RequestPageParameters, TempDistributionEntry, EmailDistributionSetupLine));
        end;
    end;

    procedure DoDistributeAdHocMail(UseDistributionCode_P: Code[10]; UseDistributionLanguageCode_P: Code[10]; UseDistributionUse4Type_P: Option; UseDistributionUse4Code_P: Code[20]; Recipient_P: Text[250]; Subject_P: Text[250]; var RecRef_P: RecordRef)
    var
        TempEmailItem_LT: Record "Email Item" temporary;
        TempDistributionEntry_LT: Record "Email Distribution Entry" temporary;
        EmailDistributionSetupLine_LT: Record "Email Distribution Setup";
        SendingDateTime_L: DateTime;
    begin
        if EmailDistributionSetupLine_LT.Get(UseDistributionCode_P, UseDistributionLanguageCode_P, UseDistributionUse4Type_P, UseDistributionUse4Code_P) then begin
            TempEmailItem_LT.ID := CreateGuid();
            TempDistributionEntry_LT."E-Mail" := Recipient_P;
            TempDistributionEntry_LT."Subject Line" := Subject_P;
            FillTempEmailItem(RecRef_P, 0, EmailDistributionSetupLine_LT, TempDistributionEntry_LT, TempEmailItem_LT, '');
            if EmailDistributionSetupLine_LT."Delay Send E-Mail" then begin
                SendingDateTime_L := GetPlannedSendingDateTime(EmailDistributionSetupLine_LT, 0DT);
                LogMailDistribution(TempEmailItem_LT, 0, SendingDateTime_L);
            end else begin
                if TempEmailItem_LT.Send(EmailDistributionSetupLine_LT."Hide Mail Dialog" or SilentDistribution) then begin
                    LogMailDistribution(TempEmailItem_LT, 1, CreateDateTime(Today(), Time()));
                end;
            end;
        end;
    end;

    local procedure DoDistributeBatchDocumentMail(RecRef: RecordRef; ReportId: Integer; DistributionCode: Code[10])
    var
        SingleDocRecRef: RecordRef;
    begin
        if RecRef.FindSet() then begin
            repeat
                SingleDocRecRef := RecRef.Duplicate();
                SingleDocRecRef.Reset();
                SingleDocRecRef.SetRecFilter();
                if DoDistributeDocumentMail(SingleDocRecRef, ReportId, DistributionCode) then begin
                    DistributionCounter += 1;
                end;
            until RecRef.Next() = 0;
        end;
    end;

    local procedure SendAsMail(RecRef: RecordRef; ReportId: Integer; EmailDistributionSetupLine: Record "Email Distribution Setup"; var TempDistributionEntry: Record "Email Distribution Entry" temporary; RequestPageParameters: Text; var ExtraFilePathList_P: DotNet List_Of_T): Boolean
    var
        TempEmailItem: Record "Email Item" temporary;
        DataTypeManagement_LC: Codeunit "Data Type Management";
        TempBlob_LT: Codeunit 4100;
        ApoFunctionsMgt_LC: Codeunit "Apo Functions Mgt.";
        FldRef_L: FieldRef;
    begin
        TempEmailItem.ID := CreateGuid();
        CopyAddtionalAttachments2EmailItemFields(TempEmailItem, ExtraFilePathList_P);
        FillTempEmailItem(RecRef, ReportId, EmailDistributionSetupLine, TempDistributionEntry, TempEmailItem, RequestPageParameters); // APO.002 SER 24.08.22
        if TempDistributionEntry."E-Mail" = '' then begin
            case EmailDistributionSetupLine."Action Empty Recipient Address" of
                EmailDistributionSetupLine."Action Empty Recipient Address"::" ":
                    exit(false);
                EmailDistributionSetupLine."Action Empty Recipient Address"::"Download Attachment":
                    exit(CreateAndDownloadPdf(RecRef, ReportId, RequestPageParameters, TempDistributionEntry, EmailDistributionSetupLine));
                EmailDistributionSetupLine."Action Empty Recipient Address"::"Print Document":
                    begin
                        Report.Print(ReportId, '', '', RecRef);
                        exit(false);
                    end;
            end;
        end;
        if TempEmailItem.Send(EmailDistributionSetupLine."Hide Mail Dialog" or SilentDistribution) then begin
            Clear(DataTypeManagement_LC);
            if RecRef.Get(RecRef.RecordId()) then begin
                if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef_L, 'Record ID') then begin
                    if RecRef.WritePermission() then begin
                        FldRef_L.Value(RecRef.RecordId());
                        RecRef.Modify(false);
                    end;
                end;
            end;
        end;
        // APO.038 JUN 26.08.24 ...
        exit(true);
    end;

    local procedure FillTempEmailItem(RecRef_P: RecordRef; ReportId_P: Integer; EmailDistributionSetupLine_PT: Record "Email Distribution Setup"; var TempDistributionEntry_PT: Record "Email Distribution Entry" temporary; var TempEmailItem_PT: Record "Email Item" temporary; RequestPageParameters_P: Text)
    var
        EmailDistributionEntry_LT: Record "Email Distribution Entry";
        DataTypeManagement_LC: Codeunit "Data Type Management";
        ApoFunctionsMgt_LC: Codeunit "Apo Functions Mgt.";
        FldRef_L: FieldRef;
        BodyText_L: Text;
        ExtraFilePaths_L: array[10] of Text;
    begin
        TempEmailItem_PT."E-Mail Distribution Code" := EmailDistributionSetupLine_PT."Email Distribution Code";
        TempEmailItem_PT."E-Mail Source Record ID" := RecRef_P.RecordId();
        TempDistributionEntry_PT."Source Record ID" := RecRef_P.RecordId();

        if (EmailDistributionSetupLine_PT."Recipient Type" = EmailDistributionSetupLine_PT."Recipient Type"::Fixed) and (TempDistributionEntry_PT."E-Mail" = '') then begin
            TempDistributionEntry_PT."E-Mail" := '';
            if EmailDistributionSetupLine_PT."Recipient Address" <> '' then begin
                TempDistributionEntry_PT."E-Mail" := EmailDistributionSetupLine_PT."Recipient Address";
            end;
        end;
        if EmailDistributionSetupLine_PT."Recipient Mailing Group" <> '' then
            TempDistributionEntry_PT."E-Mail" := AddMailAddressesFromMailingGroup(EmailDistributionSetupLine_PT."Recipient Mailing Group", TempDistributionEntry_PT."E-Mail");
        if TempDistributionEntry_PT."Sender Address" = '' then
            TryToGetFromAddress(TempEmailItem_PT, EmailDistributionSetupLine_PT)
        else
            TempEmailItem_PT."From Address" := TempDistributionEntry_PT."Sender Address";
        if TempDistributionEntry_PT."Sender Name" = '' then
            TempEmailItem_PT."From Name" := EmailDistributionSetupLine_PT."Sender Name"
        else
            TempEmailItem_PT."From Name" := TempDistributionEntry_PT."Sender Name";
        TempEmailItem_PT."Send to" := TempDistributionEntry_PT."E-Mail";
        if TempDistributionEntry_PT."Cc Recipients" = '' then
            TempEmailItem_PT."Send CC" := EmailDistributionSetupLine_PT."Cc Recipients"
        else
            TempEmailItem_PT."Send CC" := TempDistributionEntry_PT."Cc Recipients";
        if EmailDistributionSetupLine_PT."Cc Recipient Mailing Group" <> '' then
            TempEmailItem_PT."Send CC" := AddMailAddressesFromMailingGroup(EmailDistributionSetupLine_PT."Cc Recipient Mailing Group", TempEmailItem_PT."Send CC");
        if TempDistributionEntry_PT."Bcc Recipients" = '' then
            TempEmailItem_PT."Send BCC" := EmailDistributionSetupLine_PT."Bcc Recipients"
        else
            TempEmailItem_PT."Send BCC" := TempDistributionEntry_PT."Bcc Recipients";
        if EmailDistributionSetupLine_PT."Bcc Recipients Mailing Group" <> '' then
            TempEmailItem_PT."Send BCC" := AddMailAddressesFromMailingGroup(EmailDistributionSetupLine_PT."Bcc Recipients Mailing Group", TempEmailItem_PT."Send BCC");
        if TempDistributionEntry_PT."Subject Line" = '' then
            TempEmailItem_PT.Subject := GetSubjectText(EmailDistributionSetupLine_PT, TempDistributionEntry_PT)
        else
            TempEmailItem_PT.Subject := TempDistributionEntry_PT."Subject Line";
        if (not RecRef_P.IsTemporary()) and (ReportId_P > 0) then begin
            TempEmailItem_PT."Attachment File Path" := SaveReportAsPdf(RecRef_P, ReportId_P, RequestPageParameters_P);
            TempEmailItem_PT."Attachment Name" := GetDocumentFileName(RecRef_P, TempDistributionEntry_PT, EmailDistributionSetupLine_PT, '');
        end;
        TempEmailItem_PT."Language Code" := EmailDistributionSetupLine_PT."Language Code";
        TempEmailItem_PT."Use for Type" := EmailDistributionSetupLine_PT."Use for Type";
        TempEmailItem_PT."Use for Code" := EmailDistributionSetupLine_PT."Use for Code";
        // APO.026 JUN 30.11.23 ...
        // statt:
        //IF ReportId_P = 0 THEN
        TempEmailItem_PT."Plaintext Formatted" := false;
        BodyText_L := '';
        if EmailDistributionSetupLine_PT."Use Email Text from Template" then
            BodyText_L := GetMailTextBodyFromTemplate(EmailDistributionSetupLine_PT, RecRef_P, TempDistributionEntry_PT);
        if BodyText_L = '' then
            BodyText_L := ReplacePlaceholders(ApoFunctionsMgt_LC.GetTextFromBlobField(EmailDistributionSetupLine_PT, EmailDistributionSetupLine_PT.FieldNo("Text for E-Mail Body")), RecRef_P, TempDistributionEntry_PT);
        TempEmailItem_PT.SetBodyText(BodyText_L);
        if EmailDistributionEntry_LT.Get(TempDistributionEntry_PT."Entry No.") then
            TempEmailItem_PT."Email Distribution Entry No." := EmailDistributionEntry_LT."Entry No.";
        // ... APO.010 SER 24.08.22
    end;

    [TryFunction]
    local procedure TryToGetFromAddress(var TempEmailItem_PT: Record "Email Item" temporary; EmailDistributionSetupLine_PT: Record "Email Distribution Setup")
    var
        TempPossibleEmailNameValueBuffer_LT: Record "Name/Value Buffer" temporary;
        Mail_LC: Codeunit Mail;
        i: Integer;
    begin
        if EmailDistributionSetupLine_PT."Sender Address" <> '' then begin
            TempEmailItem_PT."From Address" := EmailDistributionSetupLine_PT."Sender Address";
        end else begin
            case EmailDistributionSetupLine_PT.Sender of
                EmailDistributionSetupLine_PT.Sender::"Current User":
                    begin
                        if EmailDistributionSetupLine_PT."Mail Type" = EmailDistributionSetupLine_PT."Mail Type"::SMTP then begin
                            Mail_LC.CollectCurrentUserEmailAddresses(TempPossibleEmailNameValueBuffer_LT);
                            i := 1;
                            while (i <= 4) or (TempEmailItem_PT."From Address" = '') do begin
                                case i of
                                    1:
                                        TempPossibleEmailNameValueBuffer_LT.SetFilter(Name, 'UserSetup');
                                    2:
                                        TempPossibleEmailNameValueBuffer_LT.SetFilter(Name, 'ContactEmail');
                                    3:
                                        TempPossibleEmailNameValueBuffer_LT.SetFilter(Name, 'AuthEmail');
                                    4:
                                        TempPossibleEmailNameValueBuffer_LT.SetFilter(Name, 'AD');
                                end;
                                if not TempPossibleEmailNameValueBuffer_LT.IsEmpty() then begin
                                    TempPossibleEmailNameValueBuffer_LT.FindFirst();
                                    if TempPossibleEmailNameValueBuffer_LT.Value <> '' then begin
                                        TempEmailItem_PT."From Address" := TempPossibleEmailNameValueBuffer_LT.Value;
                                        exit;
                                    end;
                                end;
                                TempPossibleEmailNameValueBuffer_LT.Reset();
                                i += 1;
                            end;
                        end;
                    end;
            end;
            if EmailDistributionSetupLine_PT."Sender Name" <> '' then
                TempEmailItem_PT."From Address" := EmailDistributionSetupLine_PT."Sender Name";
        end;
    end;

    local procedure CreateAndDownloadPdf(RecRef: RecordRef; ReportId: Integer; RequestPageParameters: Text; var TempDistributionEntry: Record "Email Distribution Entry" temporary; EmailDistributionSetupLine: Record "Email Distribution Setup"): Boolean
    var
        FileManagement: Codeunit "File Management";
        DocumentFileName: Text;
        ServerTempFileName: Text;
    begin
        ServerTempFileName := SaveReportAsPdf(RecRef, ReportId, RequestPageParameters);
        if ServerTempFileName <> '' then begin
            DocumentFileName := GetDocumentFileName(RecRef, TempDistributionEntry, EmailDistributionSetupLine, 'pdf');
            FileManagement.DownloadHandler(ServerTempFileName, SaveFileDialogTitleMsg, '', SaveFileDialogFilterMsg, DocumentFileName);
            Erase(ServerTempFileName);
            exit(true);
        end;
    end;

    local procedure SaveReportAsPdf(RecRef: RecordRef; ReportId: Integer; RequestPageParameters: Text): Text
    var
        FileManagement: Codeunit "File Management";
        ServerTempFile: File;
        OutStr: OutStream;
        ServerTempFileName: Text;
    begin
        ServerTempFileName := FileManagement.ServerTempFileName('pdf');
        ServerTempFile.Create(ServerTempFileName);
        ServerTempFile.CreateOutStream(OutStr);
        if not Report.SaveAs(ReportId, RequestPageParameters, ReportFormat::Pdf, OutStr, RecRef) then begin
            ServerTempFileName := '';
        end;
        ServerTempFile.Close();
        exit(ServerTempFileName);
    end;

    procedure LogMailDistribution(EmailItem_PT: Record "Email Item"; LogType_P: Option BeforeSending,AfterSending; SendingDateTime_P: DateTime) CreatedEntryNo_L: Integer
    var
        EmailDistributionEntry_LT: Record "Email Distribution Entry";
        TempEmailDistributionEntry_LT: Record "Email Distribution Entry" temporary;
        EmailDistributionSetup_LT: Record "Email Distribution Setup";
        ApoFunctionsMgt_LC: Codeunit "Apo Functions Mgt.";
        EmailDistrHeaderMgt_LC: Codeunit "Email Distr. Header Mgt.";
        RecRef_L: RecordRef;
        NextEntryNo_L: Integer;
    begin
        EmailDistributionEntry_LT.Reset();
        begin
            if EmailDistributionSetup_LT.Get(EmailItem_PT."E-Mail Distribution Code", EmailItem_PT."Language Code", EmailItem_PT."Use for Type", EmailItem_PT."Use for Code") then;
            Clear(EmailDistributionEntry_LT);
            EmailDistributionEntry_LT."Email Distribution Code" := EmailDistributionSetup_LT."Email Distribution Code";
            EmailDistributionEntry_LT."Language Code" := EmailDistributionSetup_LT."Language Code";
            EmailDistributionEntry_LT."Use for Type" := EmailDistributionSetup_LT."Use for Type";
            EmailDistributionEntry_LT."Use for Code" := EmailDistributionSetup_LT."Use for Code";
            EmailDistributionEntry_LT."Sender Address" := EmailItem_PT."From Address";
            if EmailItem_PT."From Name" <> '' then
                EmailDistributionEntry_LT."Sender Name" := EmailItem_PT."From Name"
            else
                EmailDistributionEntry_LT."Sender Name" := EmailItem_PT."From Address";
            EmailDistributionEntry_LT."Recipient Address" := EmailItem_PT."Send to";
            EmailDistributionEntry_LT."Cc Recipients" := EmailItem_PT."Send CC";
            // APO.010 SER 24.08.22 ...
            EmailDistributionEntry_LT."E-Mail" := EmailDistributionEntry_LT."Recipient Address";
            EmailDistributionEntry_LT."Bcc Recipients" := EmailItem_PT."Send BCC";
            // ... APO.010 SER 24.08.22
            EmailDistributionEntry_LT."Subject Line" := EmailItem_PT.Subject;
            case LogType_P of
                LogType_P::BeforeSending:
                    if (SendingDateTime_P <> 0DT) then begin
                        EmailDistributionEntry_LT."Planned Sending Date" := DT2Date(SendingDateTime_P);
                        EmailDistributionEntry_LT."Planned Sending Time" := DT2Time(SendingDateTime_P);
                        EmailDistributionEntry_LT."Planned Sending Date/Time" := SendingDateTime_P;
                        EmailDistributionEntry_LT.Status := EmailDistributionEntry_LT.Status::"Sending planned";
                    end else
                        EmailDistributionEntry_LT.Status := EmailDistributionEntry_LT.Status::New;
                LogType_P::AfterSending:
                    if (SendingDateTime_P <> 0DT) then begin
                        EmailDistributionEntry_LT."Sending Date" := DT2Date(SendingDateTime_P);
                        EmailDistributionEntry_LT."Sending Time" := DT2Time(SendingDateTime_P);
                        EmailDistributionEntry_LT.Status := EmailDistributionEntry_LT.Status::Sent;
                    end else
                        EmailDistributionEntry_LT.Status := EmailDistributionEntry_LT.Status::Error;
            end;
            EmailDistributionEntry_LT."User ID" := UserId();
            Evaluate(EmailDistributionEntry_LT."Source Record ID", Format(EmailItem_PT."E-Mail Source Record ID"));
            // APO.010 SER 24.08.22 ...
            if (EmailDistributionEntry_LT."Document Type" = '') and
                (EmailDistributionEntry_LT."Document No." = '') then
                if RecRef_L.Get(EmailItem_PT."E-Mail Source Record ID") then begin
                    EmailDistrHeaderMgt_LC.GetHeaderEntry(RecRef_L, TempEmailDistributionEntry_LT);
                    EmailDistributionEntry_LT."Document Type" := TempEmailDistributionEntry_LT."Document Type";
                    EmailDistributionEntry_LT."Document No." := TempEmailDistributionEntry_LT."Document No.";
                    EmailDistributionEntry_LT."Document Date" := TempEmailDistributionEntry_LT."Document Date";
                end;
            EmailDistributionEntry_LT.Insert();
            if (EmailDistributionEntry_LT.Status = EmailDistributionEntry_LT.Status::Sent) then
                CancelAlreadyPlanedMailDistributionWithSameRecordID(EmailDistributionEntry_LT);
            // ... APO.010 SER 24.08.22
            exit(EmailDistributionEntry_LT."Entry No.");
        end;
    end;

    local procedure CancelAlreadyPlanedMailDistributionWithSameRecordID(AlreadySentEmailDistributionEntry_PT: Record "Email Distribution Entry")
    var
        PlannedEmailDistributionEntry_LT: Record "Email Distribution Entry";
        TempPlannedEmailDistributionEntry_LT: Record "Email Distribution Entry" temporary;
        AlreadySentWithErrTxt: Label 'Blocked entry. E-Mail already sent with Entry No. %1. %2: %3';
    begin
        PlannedEmailDistributionEntry_LT.Reset();
        PlannedEmailDistributionEntry_LT.SetFilter("Entry No.", '<>%1', AlreadySentEmailDistributionEntry_PT."Entry No.");
        PlannedEmailDistributionEntry_LT.SetRange("Source Record ID", AlreadySentEmailDistributionEntry_PT."Source Record ID");
        PlannedEmailDistributionEntry_LT.SetFilter(Status, '%1|%2', PlannedEmailDistributionEntry_LT.Status::Retry, PlannedEmailDistributionEntry_LT.Status::"Sending planned");
        if PlannedEmailDistributionEntry_LT.FindSet() then begin
            repeat
                TempPlannedEmailDistributionEntry_LT := PlannedEmailDistributionEntry_LT;
                TempPlannedEmailDistributionEntry_LT.Insert();
            until PlannedEmailDistributionEntry_LT.Next() = 0;
        end;
        Clear(PlannedEmailDistributionEntry_LT);
        TempPlannedEmailDistributionEntry_LT.Reset();
        if TempPlannedEmailDistributionEntry_LT.FindSet() then begin
            repeat
                PlannedEmailDistributionEntry_LT.LockTable();
                PlannedEmailDistributionEntry_LT.Get(TempPlannedEmailDistributionEntry_LT."Entry No.");
                PlannedEmailDistributionEntry_LT.Validate(Status, PlannedEmailDistributionEntry_LT.Status::Blocked);
                PlannedEmailDistributionEntry_LT.Validate("Error Message", StrSubstNo(AlreadySentWithErrTxt, Format(AlreadySentEmailDistributionEntry_PT."Entry No."),
                                                                                                                                                                                                    PlannedEmailDistributionEntry_LT.FieldCaption("Planned Sending Date/Time"),
                                                                                                                                                                                                    PlannedEmailDistributionEntry_LT."Planned Sending Date/Time"));
                PlannedEmailDistributionEntry_LT.Validate("Planned Sending Date", 0D);
                PlannedEmailDistributionEntry_LT.Validate("Planned Sending Time", 0T);
                PlannedEmailDistributionEntry_LT.Validate("Planned Sending Date/Time", 0DT);
                PlannedEmailDistributionEntry_LT.Modify();
            until TempPlannedEmailDistributionEntry_LT.Next() = 0;
        end;
    end;

    procedure SetSendingDateTimeOnEmailDistribEntry(var EmailDistributionEntry_PT: Record "Email Distribution Entry"; SendingDate_P: Date; SendingTime_P: Time; ModifyInFunction_P: Boolean)
    begin
        EmailDistributionEntry_PT."Sending Date" := SendingDate_P;
        EmailDistributionEntry_PT."Sending Time" := SendingTime_P;
        if ModifyInFunction_P then
            EmailDistributionEntry_PT.Modify(true);
    end;

    local procedure GetMailTextBodyFromTemplate(EmailDistributionSetupLine_PT: Record "Email Distribution Setup"; RecRef_P: RecordRef; var TempDistributionEntry_PT: Record "Email Distribution Entry" temporary) BodyText_L: Text
    var
        TemplateEmailDistributionSetupLine_LT: Record "Email Distribution Setup";
        ApoFunctionsMgt_LC: Codeunit "Apo Functions Mgt.";
        RecRef_L: RecordRef;
    begin
        if RecRef_L.Get(EmailDistributionSetupLine_PT."Linked Template for Email Text") then begin
            RecRef_L.SetTable(TemplateEmailDistributionSetupLine_LT);
            BodyText_L := ReplacePlaceholders(ApoFunctionsMgt_LC.GetTextFromBlobField(TemplateEmailDistributionSetupLine_LT, TemplateEmailDistributionSetupLine_LT.FieldNo("Text for E-Mail Body")), RecRef_P, TempDistributionEntry_PT);
        end;
    end;

    procedure UploadOptionalEmailAttachments(var EmailAttachmentTmp_PT: Record "Email Attachment"; EmailItemID_P: Guid)
    var
        FileListTmp_LT: Record "Name/Value Buffer" temporary;
        ApoFunctionsMgt_LC: Codeunit "Apo Functions Mgt.";
        LineNo_L: Integer;
        TextFilterAllFiles: Label 'All Files|*.*';
    begin
        if ApoFunctionsMgt_LC.OpenFileDialogWithMultiSelect('', '', TextFilterAllFiles, FileListTmp_LT) = 0 then
            exit;
        EmailAttachmentTmp_PT.Reset();
        LineNo_L := 10000;
        if EmailAttachmentTmp_PT.FindLast() then
            LineNo_L := EmailAttachmentTmp_PT.Number + 10000;
        FileListTmp_LT.Reset();
        if FileListTmp_LT.FindSet() then begin
            repeat
                InsertTempAttachment(EmailAttachmentTmp_PT, EmailItemID_P, LineNo_L, FileListTmp_LT.Value, FileListTmp_LT.Name);
                LineNo_L += 10000;
            until FileListTmp_LT.Next() = 0;
        end;
    end;

    procedure DeleteEmailAttachmentsFromServer(TempEmailItem_PT: Record "Email Item" temporary; var TempEmailAttachment_PT: Record "Email Attachment" temporary)
    var
        FileManagement_LC: Codeunit "File Management";
    begin
        // APO.002 SER 24.08.22 ...
        FileManagement_LC.DeleteServerFile(TempEmailItem_PT."Attachment File Path");
        if FileManagement_LC.ServerFileExists(TempEmailItem_PT."Attachment File Path 2") then
            FileManagement_LC.DeleteServerFile(TempEmailItem_PT."Attachment File Path 2");
        if FileManagement_LC.ServerFileExists(TempEmailItem_PT."Attachment File Path 3") then
            FileManagement_LC.DeleteServerFile(TempEmailItem_PT."Attachment File Path 3");
        if FileManagement_LC.ServerFileExists(TempEmailItem_PT."Attachment File Path 4") then
            FileManagement_LC.DeleteServerFile(TempEmailItem_PT."Attachment File Path 4");
        if FileManagement_LC.ServerFileExists(TempEmailItem_PT."Attachment File Path 5") then
            FileManagement_LC.DeleteServerFile(TempEmailItem_PT."Attachment File Path 5");
        if FileManagement_LC.ServerFileExists(TempEmailItem_PT."Attachment File Path 6") then
            FileManagement_LC.DeleteServerFile(TempEmailItem_PT."Attachment File Path 6");
        if FileManagement_LC.ServerFileExists(TempEmailItem_PT."Attachment File Path 7") then
            FileManagement_LC.DeleteServerFile(TempEmailItem_PT."Attachment File Path 7");
        TempEmailAttachment_PT.Reset();
        if TempEmailAttachment_PT.FindSet() then
            repeat
                if FileManagement_LC.ServerFileExists(TempEmailAttachment_PT."File Path") then
                    FileManagement_LC.DeleteServerFile(TempEmailAttachment_PT."File Path");
            until TempEmailAttachment_PT.Next() = 0;
    end;

    local procedure "--- Processing ---"()
    begin
    end;

    procedure GetDistributionLine(DistributionCode: Code[10]; var TempDistributionEntry: Record "Email Distribution Entry" temporary; var EmailDistributionSetupLine: Record "Email Distribution Setup")
    begin
        if not FindDistributionLine(DistributionCode, '', EmailDistributionSetupLine."Use for Type"::Contact, TempDistributionEntry."Contact No.", EmailDistributionSetupLine) then
            if not FindDistributionLine(DistributionCode, TempDistributionEntry."Language Code", EmailDistributionSetupLine."Use for Type"::Customer, TempDistributionEntry."Customer No.", EmailDistributionSetupLine) then // APO.032 SER 26.03.24
                if not FindDistributionLine(DistributionCode, '', EmailDistributionSetupLine."Use for Type"::Customer, TempDistributionEntry."Customer No.", EmailDistributionSetupLine) then
                    if not FindDistributionLine(DistributionCode, TempDistributionEntry."Language Code", EmailDistributionSetupLine."Use for Type"::Vendor, TempDistributionEntry."Vendor No.", EmailDistributionSetupLine) then
                        if not FindDistributionLine(DistributionCode, '', EmailDistributionSetupLine."Use for Type"::Vendor, TempDistributionEntry."Vendor No.", EmailDistributionSetupLine) then
                            if not FindDistributionLine(DistributionCode, TempDistributionEntry."Language Code", EmailDistributionSetupLine."Use for Type"::"Responsibility Center", TempDistributionEntry."Resp. Center Code", EmailDistributionSetupLine) then
                                // APO.020 SER 07.08.23 ...
                                if not FindDistributionLine(DistributionCode, TempDistributionEntry."Language Code", EmailDistributionSetupLine."Use for Type"::"Location Code", TempDistributionEntry."Location Code", EmailDistributionSetupLine) then
                                    if not FindDistributionLine(DistributionCode, TempDistributionEntry."Language Code", EmailDistributionSetupLine."Use for Type"::"Internal Company", TempDistributionEntry."Internal Company", EmailDistributionSetupLine) then
                                        if not FindDistributionLine(DistributionCode, TempDistributionEntry."Language Code", EmailDistributionSetupLine."Use for Type"::" ", '', EmailDistributionSetupLine) then
                                            if not FindDistributionLine(DistributionCode, '', EmailDistributionSetupLine."Use for Type"::" ", '', EmailDistributionSetupLine) then begin
                                                EmailDistributionSetupLine.SetRecFilter();
                                                EmailDistributionSetupLine.FindFirst(); // error
                                            end;
        // ... APO.020 SER 07.08.23
        // APO.032 SER 26.03.24
    end;

    [TryFunction]
    procedure TrySearchReportSelection(HeaderDoc: Variant; var ReportId: Integer; var DistributionCode: Code[10])
    begin
        SearchReportSelection(HeaderDoc, ReportId, DistributionCode);
    end;

    local procedure SearchReportSelection(HeaderDoc: Variant; var ReportId: Integer; var DistributionCode: Code[10])
    var
        SalesHeader: Record "Sales Header";
        PurchHeader: Record "Purchase Header";
        ReportSelections: Record "Report Selections";
        ServHeader: Record "Service Header";
        ServiceContractHeader: Record "Service Contract Header";
        DataTypeMgt: Codeunit "Data Type Management";
        RecRef: RecordRef;
    begin
        DataTypeMgt.GetRecordRef(HeaderDoc, RecRef);
        ReportSelections.Reset();
        ReportSelections.SetFilter("Report ID", '<>0');
        case RecRef.Number() of
            Database::"Sales Header":
                begin
                    RecRef.SetTable(SalesHeader);

                    case SalesHeader."Document Type" of

                        SalesHeader."Document Type"::Quote:
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"S.Quote");
                            end;

                        SalesHeader."Document Type"::Order:
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"S.Order");
                            end;

                        SalesHeader."Document Type"::"Blanket Order":
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"S.Blanket");
                            end;

                        SalesHeader."Document Type"::"Return Order":
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"S.Return");
                            end;

                        else
                            Error(StrSubstNo(TextTypeNotSupported, Format(SalesHeader."Document Type"), RecRef.Name()));

                    end;
                end;

            Database::"Sales Invoice Header":
                begin
                    ReportSelections.SetRange(Usage, ReportSelections.Usage::"S.Invoice");
                end;

            Database::"Sales Cr.Memo Header":
                begin
                    ReportSelections.SetRange(Usage, ReportSelections.Usage::"S.Cr.Memo");
                end;

            Database::"Sales Shipment Header":
                begin
                    ReportSelections.SetRange(Usage, ReportSelections.Usage::"S.Shipment");
                end;

            Database::"Purchase Header":
                begin
                    RecRef.SetTable(PurchHeader);

                    case PurchHeader."Document Type" of

                        PurchHeader."Document Type"::Quote:
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"P.Quote");
                            end;

                        PurchHeader."Document Type"::Order:
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"P.Order");
                            end;

                        PurchHeader."Document Type"::"Blanket Order":
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"P.Blanket");
                            end;

                        PurchHeader."Document Type"::"Return Order":
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"P.Return");
                            end;

                        else
                            Error(StrSubstNo(TextTypeNotSupported, Format(PurchHeader."Document Type"), RecRef.Name()));

                    end;
                end;

            Database::"Purch. Cr. Memo Hdr.":
                begin
                    ReportSelections.SetRange(Usage, ReportSelections.Usage::"P.Cr.Memo");
                end;

            Database::"Return Shipment Header":
                begin
                    ReportSelections.SetRange(Usage, ReportSelections.Usage::"P.Ret.Shpt.");
                end;

            Database::"Service Header":
                begin
                    RecRef.SetTable(ServHeader);

                    case ServHeader."Document Type" of

                        ServHeader."Document Type"::Quote:
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"SM.Quote");
                            end;

                        ServHeader."Document Type"::Order:
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"SM.Order");
                            end;

                        else
                            Error(StrSubstNo(TextTypeNotSupported, Format(ServHeader."Document Type"), RecRef.Name()));

                    end;
                end;

            Database::"Service Shipment Header":
                begin
                    ReportSelections.SetRange(Usage, ReportSelections.Usage::"SM.Shipment");
                end;

            Database::"Service Invoice Header":
                begin
                    ReportSelections.SetRange(Usage, ReportSelections.Usage::"SM.Invoice");
                end;

            Database::"Service Cr.Memo Header":
                begin
                    ReportSelections.SetRange(Usage, ReportSelections.Usage::"SM.Credit Memo");
                end;

            Database::"Service Contract Header":
                begin
                    RecRef.SetTable(ServiceContractHeader);
                    case ServiceContractHeader."Contract Type" of

                        ServiceContractHeader."Contract Type"::Quote:
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"SM.Contract Quote");
                            end;

                        ServiceContractHeader."Contract Type"::Contract:
                            begin
                                ReportSelections.SetRange(Usage, ReportSelections.Usage::"SM.Contract");
                            end;

                        else
                            Error(StrSubstNo(TextTypeNotSupported, Format(ServiceContractHeader."Contract Type"), RecRef.Name()));

                    end;
                end;

            Database::"Issued Reminder Header":
                begin
                    ReportSelections.SetRange(Usage, ReportSelections.Usage::Reminder);
                end;

            Database::"Issued Fin. Charge Memo Header":
                begin
                    ReportSelections.SetRange(Usage, ReportSelections.Usage::"Fin.Charge");
                end;
            else
                Error(StrSubstNo(TextTypeNotSupported, RecRef.Name(), RecRef.Name()));

        end;

        ReportSelections.FindFirst();
        ReportId := ReportSelections."Report ID";
        DistributionCode := ReportSelections."Email Distribution Code";
    end;

    local procedure GetDistributionType(RecRef: RecordRef; EmailDistributionSetupLine: Record "Email Distribution Setup"): Integer
    var
        TempDistributionEntry: Record "Email Distribution Entry" temporary;
        DistributionEntryRecRef: RecordRef;
        FldRef: FieldRef;
        Result: Integer;
    begin
        if SilentDistribution then begin
            exit(TempDistributionEntry."Distribution Type"::Mail);
        end;

        if not EmailDistributionSetupLine."Use Distribution Selection" then begin
            exit(TempDistributionEntry."Distribution Type"::Mail);
        end;

        DistributionEntryRecRef.GetTable(TempDistributionEntry);
        FldRef := DistributionEntryRecRef.Field(TempDistributionEntry.FieldNo(TempDistributionEntry."Distribution Type"));
        Result := StrMenu(FldRef.OptionCaption());

        if Result = 0 then begin
            Error('');
        end;

        exit(Result - 1);
    end;

    local procedure RunAndGetRequestPage(ReportId: Integer): Text
    var
        RequestParameters: Text;
    begin
        RequestParameters := Report.RunRequestPage(ReportId);
        if RequestParameters = '' then begin
            Error('');
        end;

        exit(RequestParameters);
    end;

    local procedure GetSubjectText(EmailDistributionSetupLine: Record "Email Distribution Setup"; var TempDistributionEntry: Record "Email Distribution Entry" temporary): Text
    var
        Vendor_LT: Record Vendor;
        CompanyInformation: Record "Company Information";
        RecRef: RecordRef;
        EmailSubject: Text;
    begin
        Clear(RecRef);
        if RecRef.Get(TempDistributionEntry."Source Record ID") then;
        case EmailDistributionSetupLine."Subject Line Format" of
            EmailDistributionSetupLine."Subject Line Format"::"Recipient Name - Source Type - Source No.":
                begin
                    exit(StrSubstNo('%1 - %2 %3', TempDistributionEntry."To Name", TempDistributionEntry."Document Type", TempDistributionEntry."Document No."));
                end;
            EmailDistributionSetupLine."Subject Line Format"::"Company Name - Source Type - Source No.":
                begin
                    if TempDistributionEntry."Vendor No." <> '' then begin
                        Vendor_LT.Get(TempDistributionEntry."Vendor No.");
                        if Vendor_LT."Internal Company" <> '' then begin
                            if CompanyInformation.Get(Vendor_LT."Internal Company") then
                                exit(StrSubstNo('%1 - %2 %3', CompanyInformation.Name, TempDistributionEntry."Document Type", TempDistributionEntry."Document No."));
                        end;
                    end;
                    if CompanyInformation.Get() then begin
                        exit(StrSubstNo('%1 - %2 %3', CompanyInformation.Name, TempDistributionEntry."Document Type", TempDistributionEntry."Document No."));
                    end;
                    Clear(CompanyInformation);
                end;

            EmailDistributionSetupLine."Subject Line Format"::"Subject Text":
                begin
                    EmailSubject := EmailDistributionSetupLine."Subject Text";
                    EmailSubject := ReplacePlaceholders(EmailSubject, RecRef, TempDistributionEntry);
                    EmailSubject := StrSubstNo(EmailSubject, GTS[1], GTS[2], GTS[3], GTS[4], GTS[5], GTS[6], GTS[7], GTS[8], GTS[9], GTS[10]);
                    EmailDistributionSetupLine."Subject Text" := EmailSubject;
                    exit(EmailDistributionSetupLine."Subject Text");
                end;

        end;
    end;

    procedure ReplacePlaceholders(BodyText: DotNet String; RecRef: RecordRef; var TempDistributionEntry: Record "Email Distribution Entry" temporary): Text
    var
        GroupCollection: DotNet GroupCollection;
        Match: DotNet Match;
        Regex: DotNet Regex;
        ValueString: DotNet String;
        VariableName: Text;
        i: Integer;
    begin
        ValueString := BodyText;
        if ValueString.Contains('[') then begin
            Match := Regex.Match(Format(ValueString), '\[([^]]*)\]');
            while Match.Success do begin
                GroupCollection := Match.Groups;

                VariableName := Format(GroupCollection.Item(1));
                if Match.Length = 1 then
                    exit;
                ValueString := ValueString.Replace(StrSubstNo('[%1]', VariableName), GetPlaceholderKeyValue(RecRef, TempDistributionEntry, VariableName));

                Match := Match.NextMatch;
                i += 1;
                if i = 100 then Error(TextRegExInfiniteLoop);
            end;
        end;
        exit(Format(ValueString));
    end;

    local procedure GetPlaceholderKeyValue(RecRef: RecordRef; var TempDistributionEntry: Record "Email Distribution Entry" temporary; VariableName: Text): Text
    var
        SalesPerson: Record "Salesperson/Purchaser";
        CompanyInfo: Record "Company Information";
        UserSetup: Record "User Setup";
        GeneralLedgerSetup: Record "General Ledger Setup";
        ExtendedTextHeader_LT: Record "Extended Text Header";
        ExtendedTextLineTmp_LT: Record "Extended Text Line" temporary;
        ShippingAgent_LT: Record "Shipping Agent";
        ReminderLevel_LT: Record "Reminder Level";
        NameValueBuffer_LTT: Record "Name/Value Buffer" temporary;
        Contact: Record Contact;
        ResponsibilityCenter_LT: Record "Responsibility Center";
        ShippingAgentServices_LT: Record "Shipping Agent Services";
        ApoSetup_LT: Record "Apo Setup";
        TransferExtendedText_LC: Codeunit "Transfer Extended Text";
        DataTypeManagement_LC: Codeunit "Data Type Management";
        EmailDistributionHeaderMgt: Codeunit "Email Distr. Header Mgt.";
        FldRef: FieldRef;
        FldRef2: FieldRef;
        Date1_L: Date;
        CurrencyCode: Text;
        Text1_L: Text;
        Text2_L: Text;
        Code1_L: Code[20];
        Code2_L: Code[20];
        Cr_L: Char;
        Lf_L: Char;
        AmountInclVat: Decimal;
        PmtDiscountPercent: Decimal;
        Counter_L: Integer;
    begin
        Cr_L := 13;
        Lf_L := 10;

        case VariableName of
            'SalesPerson.Name',
            // APO.021 SER 08.08.23 ...
            'SalesPerson.PhoneNo',
            'SalesPerson.E-Mail',
            // ... APO.021 SER 08.08.23
            'SalesPerson.JobTitle':
                begin
                    if not SalesPerson.Get(EmailDistributionHeaderMgt.GetFirstSourceTableFieldCode(RecRef, Database::"Salesperson/Purchaser")) then
                        if UserSetup.Get(UserId()) then
                            if SalesPerson.Get(UserSetup."Salespers./Purch. Code") then;
                    if SalesPerson.Code <> '' then
                        case VariableName of
                            'SalesPerson.Name':
                                exit(SalesPerson.Name);
                            'SalesPerson.JobTitle':
                                exit(SalesPerson."Job Title");
                            // APO.021 SER 08.08.23 ...
                            'SalesPerson.PhoneNo':
                                exit(SalesPerson."Phone No.");
                            'SalesPerson.E-Mail':
                                exit(SalesPerson."E-Mail");
                        // ... APO.021 SER 08.08.23
                        end;
                end;
            // APO.021 SER 08.08.23 ...
            'CompanyInfo.Name', 'CompanyInfo.Name2', 'CompanyInfo.Address', 'CompanyInfo.Address2', 'CompanyInfo.PostCode', 'CompanyInfo.City',
            'CompanyInfo.VATRegistrationNo', 'CompanyInfo.CountryRegionCode', 'CompanyInfo.MailSignatureText':
                begin
                    if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef, 'Internal Company') then begin
                        if Format(FldRef.Value()) <> '' then begin
                            if not CompanyInfo.Get(FldRef.Value()) then
                                exit('');
                        end else
                            if not CompanyInfo.Get() then
                                exit('');
                    end else begin
                        if not CompanyInfo.Get() then
                            exit('');
                    end;
                    case VariableName of
                        'CompanyInfo.Name':
                            exit(CompanyInfo.Name);
                        'CompanyInfo.Name2':
                            exit(CompanyInfo."Name 2");
                        'CompanyInfo.Address':
                            exit(CompanyInfo.Address);
                        'CompanyInfo.Address2':
                            exit(CompanyInfo."Address 2");
                        'CompanyInfo.PostCode':
                            exit(CompanyInfo."Post Code");
                        'CompanyInfo.City':
                            exit(CompanyInfo.City);
                        'CompanyInfo.CountryRegionCode':
                            exit(CompanyInfo."Country/Region Code");
                        'CompanyInfo.VATRegistrationNo':
                            exit(CompanyInfo."VAT Registration No.");
                        'CompanyInfo.MailSignatureText':
                            if CompanyInfo."E-Mail Signature Text Code" <> '' then begin
                                Code1_L := '';
                                if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef, 'Language Code') then
                                    Code1_L := FldRef.Value();
                                Date1_L := 0D;
                                if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef2, 'Document Date') then
                                    Date1_L := FldRef2.Value();
                                ExtendedTextHeader_LT.SetRange("No.", CompanyInfo."E-Mail Signature Text Code");
                                ExtendedTextHeader_LT.SetRange("Table Name", ExtendedTextHeader_LT."Table Name"::"Standard Text");
                                if TransferExtendedText_LC.ReadLines(ExtendedTextHeader_LT, Date1_L, Code1_L) then begin
                                    TransferExtendedText_LC.GetTempExtTextLine(ExtendedTextLineTmp_LT);
                                    ExtendedTextLineTmp_LT.Reset();
                                    if ExtendedTextLineTmp_LT.FindSet() then begin
                                        Cr_L := 13;
                                        Lf_L := 10;
                                        Text1_L := '';
                                        repeat
                                            Text1_L += ExtendedTextLineTmp_LT."Ext. Report Text" + Format(Cr_L) + Format(Lf_L); // new line for each text line
                                        until ExtendedTextLineTmp_LT.Next() = 0;
                                    end;
                                    exit(Text1_L);
                                end;
                            end;
                    end;
                end;
            // 'CompanyInfo.Name':
            //  IF CompanyInfo.GET THEN
            //    EXIT(CompanyInfo.Name);
            // 'CompanyInfo.Address':
            //  IF CompanyInfo.GET THEN
            //    EXIT(CompanyInfo.Address);
            // 'CompanyInfo.PostCode':
            //  IF CompanyInfo.GET THEN
            //    EXIT(CompanyInfo."Post Code");
            // 'CompanyInfo.City':
            //  IF CompanyInfo.GET THEN
            //    EXIT(CompanyInfo.City);
            // ... APO.021 SER 08.08.23
            'DocumentType':
                exit(TempDistributionEntry."Document Type");
            'DocumentNo':
                exit(TempDistributionEntry."Document No.");
            'DocumentDate':
                exit(Format(TempDistributionEntry."Document Date"));
            'EmailAddress':
                exit(TempDistributionEntry."E-Mail");
            'FaxNo':
                exit(TempDistributionEntry."Fax No.");
            'LanguageCode':
                exit(TempDistributionEntry."Language Code");
            'CustomerNo':
                exit(TempDistributionEntry."Customer No.");
            'VendorNo':
                exit(TempDistributionEntry."Vendor No.");
            'ContactNo':
                exit(TempDistributionEntry."Contact No.");
            'RespCenterCode':
                exit(TempDistributionEntry."Resp. Center Code");
            // APO.004 SER 08.08.23 ...
            // APO.029 JUN 17.01.24 ...
            'RespCenter.Name', 'RespCenter.Name2', 'RespCenter.Address', 'RespCenter.Address2', 'RespCenter.City', 'RespCenter.PostCode', 'RespCenter.PhoneNo', 'RespCenter.E-Mail', 'RespCenter.MailSignatureText':
                // statt:
                //'RespCenter.Name','RespCenter.Name2','RespCenter.Address','RespCenter.Address2','RespCenter.City','RespCenter.PostCode','RespCenter.PhoneNo','RespCenter.E-Mail':
                // ... APO.029 JUN 17.01.24
                begin
                    if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef, 'Responsibility Center') then begin
                        if ResponsibilityCenter_LT.Get(FldRef.Value()) then begin
                            case VariableName of
                                'RespCenter.Name':
                                    exit(ResponsibilityCenter_LT.Name);
                                'RespCenter.Name2':
                                    exit(ResponsibilityCenter_LT."Name 2");
                                'RespCenter.Address':
                                    exit(ResponsibilityCenter_LT.Address);
                                'RespCenter.Address2':
                                    exit(ResponsibilityCenter_LT."Address 2");
                                'RespCenter.City':
                                    exit(ResponsibilityCenter_LT.City);
                                'RespCenter.PostCode':
                                    exit(ResponsibilityCenter_LT."Post Code");
                                'RespCenter.PhoneNo':
                                    exit(ResponsibilityCenter_LT."Phone No.");
                                'RespCenter.E-Mail':
                                    exit(ResponsibilityCenter_LT."E-Mail");
                                // APO.029 JUN 17.01.24 ...
                                'RespCenter.MailSignatureText':
                                    begin
                                        Code1_L := '';
                                        if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef, 'Language Code') then
                                            Code1_L := FldRef.Value();
                                        Date1_L := 0D;
                                        if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef2, 'Document Date') then
                                            Date1_L := FldRef2.Value();
                                        ExtendedTextHeader_LT.SetRange("No.", ResponsibilityCenter_LT."E-Mail Signature Text Code");
                                        ExtendedTextHeader_LT.SetRange("Table Name", ExtendedTextHeader_LT."Table Name"::"Standard Text");
                                        if TransferExtendedText_LC.ReadLines(ExtendedTextHeader_LT, Date1_L, Code1_L) then begin
                                            TransferExtendedText_LC.GetTempExtTextLine(ExtendedTextLineTmp_LT);
                                            ExtendedTextLineTmp_LT.Reset();
                                            if ExtendedTextLineTmp_LT.FindSet() then begin
                                                Text1_L := '';
                                                repeat
                                                    Text1_L += ExtendedTextLineTmp_LT."Ext. Report Text";
                                                    if ExtendedTextLineTmp_LT."Print Newline after Text" then
                                                        Text1_L += Format(Cr_L) + Format(Lf_L); // new line for each text line
                                                until ExtendedTextLineTmp_LT.Next() = 0;
                                            end;
                                            exit(Text1_L);
                                        end;
                                    end;
                            // ... APO.029 JUN 17.01.24
                            end;
                        end;
                    end;
                end;
            // ... APO.004 SER 08.08.23
            'ItemNo':
                exit(TempDistributionEntry."Item No.");
            'AccountNo':
                exit(TempDistributionEntry."Account No.");
            'TransferCode':
                exit(TempDistributionEntry."Transfer Code");
            'ToName':
                exit(TempDistributionEntry."To Name");
            'DueDate':
                if RecRef.FieldExist(24) then begin
                    FldRef := RecRef.Field(24);
                    exit(Format(FldRef.Value()));
                end;
            'PmtDiscountPercent':
                if RecRef.FieldExist(25) then begin
                    FldRef := RecRef.Field(25);
                    exit(Format(FldRef.Value()));
                end;
            'PmtDiscountDate':
                if RecRef.FieldExist(26) then begin
                    FldRef := RecRef.Field(26);
                    exit(Format(FldRef.Value()));
                end;
            'AmountInclVat':
                if RecRef.FieldExist(61) then begin
                    FldRef := RecRef.Field(61);
                    FldRef.CalcField();
                    exit(Format(FldRef.Value()));
                end;
            'PmtDiscountAmount',
            'Amount-PmtDiscountAmount':
                if RecRef.FieldExist(25) and RecRef.FieldExist(61) then begin
                    FldRef := RecRef.Field(25);
                    PmtDiscountPercent := FldRef.Value();
                    FldRef := RecRef.Field(61);
                    FldRef.CalcField();
                    AmountInclVat := FldRef.Value();
                    if (PmtDiscountPercent <> 0) and (AmountInclVat <> 0) then
                        case VariableName of
                            'PmtDiscountAmount':
                                exit(Format(AmountInclVat * (PmtDiscountPercent / 100)));
                            'Amount-PmtDiscountAmount':
                                exit(Format(AmountInclVat * ((100 - PmtDiscountPercent) / 100)));
                        end;
                end;
            'Currency':
                begin
                    CurrencyCode := EmailDistributionHeaderMgt.GetFirstSourceTableFieldCode(RecRef, Database::Currency);
                    if CurrencyCode <> '' then
                        exit(CurrencyCode);
                    GeneralLedgerSetup.Get();
                    exit(GeneralLedgerSetup."LCY Code");
                end;
            'Salutation':
                if Contact.Get(TempDistributionEntry."Contact No.") then
                    exit(Contact.GetSalutation(0, TempDistributionEntry."Language Code"));
            'ExternalDocumentNo':
                begin
                    if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef, 'External Document No.') then
                        exit(Format(FldRef.Value()));
                end;
            'OrderDate':
                begin
                    if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef, 'Order Date') then
                        exit(Format(FldRef.Value()));
                end;
            'DetailsFromSendingRecord':
                begin
                    exit(GetDetailsFromSendingRecord(RecRef));
                end;
            // APO.012 JUN 26.10.22 ...
            'ShippingAgent':
                begin
                    if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef, 'Shipping Agent Code') then begin
                        ShippingAgent_LT.Get(FldRef.Value());
                        exit(ShippingAgent_LT.Name);
                    end;
                end;
            // ... APO.012 JUN 26.10.22
            // APO.019 JUN 07.08.23 ...
            'ReminderLevelText':
                begin
                    if RecRef.Number() <> Database::"Issued Reminder Header" then
                        exit('');
                    if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef, 'Reminder Terms Code') then
                        if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef2, 'Reminder Level') then
                            if ReminderLevel_LT.Get(FldRef.Value(), FldRef2.Value()) then
                                if ReminderLevel_LT."E-Mail Text Code" <> '' then begin
                                    DataTypeManagement_LC.FindFieldByName(RecRef, FldRef, 'Language Code');
                                    Code1_L := FldRef.Value();
                                    DataTypeManagement_LC.FindFieldByName(RecRef, FldRef2, 'Posting Date');
                                    Date1_L := FldRef2.Value();
                                    ExtendedTextHeader_LT.SetRange("No.", ReminderLevel_LT."E-Mail Text Code");
                                    ExtendedTextHeader_LT.SetRange("Table Name", ExtendedTextHeader_LT."Table Name"::"Standard Text");
                                    ExtendedTextHeader_LT.SetRange(Reminder, true);
                                    if TransferExtendedText_LC.ReadLines(ExtendedTextHeader_LT, Date1_L, Code1_L) then begin
                                        TransferExtendedText_LC.GetTempExtTextLine(ExtendedTextLineTmp_LT);
                                        ExtendedTextLineTmp_LT.Reset();
                                        if ExtendedTextLineTmp_LT.FindSet() then begin
                                            Text1_L := '';
                                            repeat
                                                Text1_L += ExtendedTextLineTmp_LT.Text;
                                                // APO.025 JUN 27.11.23 ...
                                                if ExtendedTextLineTmp_LT."Print Newline after Text" then
                                                    Text1_L += Format(Cr_L) + Format(Lf_L);
                                            // ... APO.025 JUN 27.11.23
                                            until ExtendedTextLineTmp_LT.Next() = 0;
                                        end;
                                        exit(Text1_L);
                                    end;
                                end;
                end;
            // ... APO.019 JUN 07.08.23
            // APO.039 JUN 03.09.24 ...
            'ReminderSubjectLine':
                begin
                    if RecRef.Number() <> Database::"Issued Reminder Header" then
                        exit('');
                    if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef, 'Reminder Terms Code') then
                        if DataTypeManagement_LC.FindFieldByName(RecRef, FldRef2, 'Reminder Level') then
                            if ReminderLevel_LT.Get(FldRef.Value(), FldRef2.Value()) then
                                if ReminderLevel_LT."Email Subject Text Code" <> '' then begin
                                    DataTypeManagement_LC.FindFieldByName(RecRef, FldRef, 'Language Code');
                                    Code1_L := FldRef.Value();
                                    DataTypeManagement_LC.FindFieldByName(RecRef, FldRef2, 'Posting Date');
                                    Date1_L := FldRef2.Value();
                                    ExtendedTextHeader_LT.SetRange("No.", ReminderLevel_LT."Email Subject Text Code");
                                    ExtendedTextHeader_LT.SetRange("Table Name", ExtendedTextHeader_LT."Table Name"::"Standard Text");
                                    ExtendedTextHeader_LT.SetRange(Reminder, true);
                                    if TransferExtendedText_LC.ReadLines(ExtendedTextHeader_LT, Date1_L, Code1_L) then begin
                                        TransferExtendedText_LC.GetTempExtTextLine(ExtendedTextLineTmp_LT);
                                        ExtendedTextLineTmp_LT.Reset();
                                        if ExtendedTextLineTmp_LT.FindFirst() then
                                            // APO.040 JUN 04.09.24 ...
                                            Text1_L := StrSubstNo('%1 - %2 %3', TempDistributionEntry."To Name", ExtendedTextLineTmp_LT.Text, TempDistributionEntry."Document No.");
                                        // statt:
                                        //Text1_L := STRSUBSTNO('%1 - %2. %3 %4',TempDistributionEntry."To Name", ReminderLevel_LT."No.",ExtendedTextLineTmp_LT.Text,TempDistributionEntry."Document No.");
                                        // ... APO.040 JUN 04.09.24
                                        exit(Text1_L);
                                    end;
                                end;
                end;
            // ... APO.039 JUN 03.09.24
            else
                exit(VariableName);
        end;

        // >> cc|mail
        // 01 SalesPerson.Name
        // 02 SalesPerson."Job Title"
        // 03 CompanyInfo.Name;
        // 04 CompanyInfo.Address;
        // 05 CompanyInfo."Post Code";
        // 06 CompanyInfo.City;
        // 07 Document Type [+ Document No]
        // 10 DueDate
        // 11 Pmt. Discount Date
        // 12 Pmt. Discount %
        // 13 LCY Code
        // 14 Amount Inlc Vat
        // 15 Payment Discount Amount
        // 16 Amount - Payment Discount Amount
        // 20 GetSalutation (Language)
        // 21 ToAddress
        // << cc|mail
    end;

    local procedure GetDetailsFromSendingRecord(RecRef_P: RecordRef) DetailsFromRecord_L: Text
    var
        EmailDistrRecordCollLines_LT: Record "Email Distr. Record Coll.";
        DataTypeManagement_LC: Codeunit "Data Type Management";
        RecRef_L: RecordRef;
        FldRef_L: FieldRef;
        CurrPrimaryKeyFieldValuesAsText_L: Text;
        LastPrimaryKeyFieldValuesAsText_L: Text;
        EntryNoFilter_L: Integer;
    begin
        DetailsFromRecord_L := '';
        case RecRef_P.Number() of
            Database::"Email Distr. Record Coll.":
                begin
                    if DataTypeManagement_LC.GetRecordRefAndFieldRef(RecRef_P, EmailDistrRecordCollLines_LT.FieldNo("No."), RecRef_P, FldRef_L) then begin
                        EmailDistrRecordCollLines_LT.SetRange(Type, EmailDistrRecordCollLines_LT.Type::"Collection Line");
                        EntryNoFilter_L := FldRef_L.Value();
                        EmailDistrRecordCollLines_LT.SetRange("No.", EntryNoFilter_L);
                        if EmailDistrRecordCollLines_LT.FindSet() then begin
                            repeat
                                EmailDistrRecordCollLines_LT.CalcFields("Field Name");
                                if DetailsFromRecord_L <> '' then
                                    DetailsFromRecord_L += '<br>';
                                if RecRef_L.Get(EmailDistrRecordCollLines_LT."Record ID") then
                                    DetailsFromRecord_L += StrSubstNo('<b>%1</b>%2%3', GetPrimaryKeyFieldValuesAsText(RecRef_L), '<br>',
                                                                                                                          StrSubstNo('%1 %2: %3',
                                                                                                                              EmailDistrRecordCollLines_LT."Field Name",
                                                                                                                              EmailDistrRecordCollLines_LT.FieldCaption("New Value"),
                                                                                                                              EmailDistrRecordCollLines_LT."New Value"));
                            until EmailDistrRecordCollLines_LT.Next() = 0;
                        end;
                    end;
                end;
        end;
    end;

    local procedure "--- Help Fcts. ---"()
    begin
    end;

    local procedure CleanUpEMailAttachments(var EmailDistributionEntry_PT: Record "Email Distribution Entry")
    var
        EmailDistributionAttachment_LT: Record "Email Distribution Attachment";
    begin
        EmailDistributionAttachment_LT.Reset();
        EmailDistributionAttachment_LT.SetRange("Email Distribution Code", EmailDistributionEntry_PT."Email Distribution Code");
        EmailDistributionAttachment_LT.SetRange("Language Code", EmailDistributionEntry_PT."Language Code");
        //EmailDistributionAttachment_LT.SETRANGE("Use for Type", EmailDistributionEntry_PT."Use for Type");
        //EmailDistributionAttachment_LT.SETRANGE("Use for Code", EmailDistributionEntry_PT."Use for Code");
        EmailDistributionAttachment_LT.SetRange("Temporary", true);
        EmailDistributionAttachment_LT.DeleteAll();
    end;

    local procedure GetFirstSourceTableFieldCode(RecordVariant: Variant; RelationTableNo: Integer): Text
    var
        FldRec: Record "Field";
        DataTypeMgt: Codeunit "Data Type Management";
        RecRef: RecordRef;
        FldRef: FieldRef;
    begin
        DataTypeMgt.GetRecordRef(RecordVariant, RecRef);

        FldRec.SetRange(TableNo, RecRef.Number());
        FldRec.SetRange(Type, FldRec.Type::Code);
        FldRec.SetRange(Class, FldRec.Class::Normal);
        FldRec.SetRange(Enabled, true);
        FldRec.SetRange(RelationTableNo, RelationTableNo);

        if FldRec.FindFirst() then begin
            FldRef := RecRef.Field(FldRec."No.");

            exit(FldRef.Value());
        end;
    end;

    local procedure FindDistributionLine(DistributionCode: Code[10]; LanguageCode: Code[10]; UseForType: Integer; UseForCode: Code[20]; var EmailDistributionSetupLine: Record "Email Distribution Setup"): Boolean
    begin
        EmailDistributionSetupLine.Reset();
        EmailDistributionSetupLine.SetRange("Email Distribution Code", DistributionCode);
        EmailDistributionSetupLine.SetRange("Language Code", LanguageCode);
        EmailDistributionSetupLine.SetRange("Use for Type", UseForType);
        EmailDistributionSetupLine.SetRange("Use for Code", UseForCode);

        exit(EmailDistributionSetupLine.FindFirst());
    end;

    procedure GetDocumentFileName(RecRef: RecordRef; var TempDistributionEntry: Record "Email Distribution Entry" temporary; EmailDistributionSetupLine: Record "Email Distribution Setup"; FileFormat_P: Text) FileNameText_L: Text
    var
        FileMgt: Codeunit "File Management";
        TextField: array[3] of Text;
        TextAttntFileNameWithFormat_L: Label '%1 - %2 - %3.%4', Comment = '%1 = Company Name. %2 = Document Type %3 = Invoice No.';
    begin
        case EmailDistributionSetupLine."Attached File Name Format" of
            EmailDistributionSetupLine."Attached File Name Format"::"Company Name - Document Type - Document No":
                if FileFormat_P = '' then begin
                    exit(StrSubstNo(TextAttachmentFileName, FileMgt.StripNotsupportChrInFileName(GetCompanyName(RecRef)), TempDistributionEntry."Document Type", TempDistributionEntry."Document No."));
                end else begin
                    exit(StrSubstNo(TextAttntFileNameWithFormat_L, FileMgt.StripNotsupportChrInFileName(GetCompanyName(RecRef)), TempDistributionEntry."Document Type", TempDistributionEntry."Document No.", FileFormat_P));
                end;
            EmailDistributionSetupLine."Attached File Name Format"::"Purchase No - Vendor - Order Date":
                begin
                    GetRecFieldsForPurchaseHeader(RecRef, TextField);
                    if FileFormat_P = '' then begin
                        exit(StrSubstNo(TextAttachmentFileName, TextField[1], TextField[2], TextField[3]));
                    end else begin
                        exit(StrSubstNo(TextAttntFileNameWithFormat_L, TextField[1], TextField[2], TextField[3], FileFormat_P));
                    end;
                end;
            EmailDistributionSetupLine."Attached File Name Format"::"Subject Text":
                begin
                    FileNameText_L := EmailDistributionSetupLine."Attached File Name";
                    FileNameText_L := ReplacePlaceholders(FileNameText_L, RecRef, TempDistributionEntry);
                    FileNameText_L := StrSubstNo(FileNameText_L, GTS[1], GTS[2], GTS[3], GTS[4], GTS[5], GTS[6], GTS[7], GTS[8], GTS[9], GTS[10]);
                    if FileFormat_P = '' then begin
                        exit(StrSubstNo('%1.%2', FileNameText_L, 'pdf'));
                    end else begin
                        exit(StrSubstNo('%1.%2', FileNameText_L, FileFormat_P));
                    end;
                end;
            // APO.027 JUN 04.12.23 ...
            EmailDistributionSetupLine."Attached File Name Format"::"Document Type - Document No":
                begin
                    if FileFormat_P = '' then
                        exit(StrSubstNo('%1 - %2.%3', TempDistributionEntry."Document Type", TempDistributionEntry."Document No.", 'pdf'))
                    else
                        exit(StrSubstNo('%1 - %2.%3', TempDistributionEntry."Document Type", TempDistributionEntry."Document No.", FileFormat_P));
                end;
        // ... APO.027 JUN 04.12.23
        end;
    end;

    procedure MultiplyDistributionLine(SrcDistributionLine_PT: Record "Email Distribution Setup")
    var
        SelectedTarget_L: Integer;
        TargetStringMenuQstTxt: Label 'Customer,Vendor,Contact,Responsibility Center,Internal Usage,Template,Location Code,Internal Company';
    begin
        if SrcDistributionLine_PT."Use for Type" = SrcDistributionLine_PT."Use for Type"::" " then begin
            SrcDistributionLine_PT.TestField(SrcDistributionLine_PT."Use for Type");
        end else begin
            SelectedTarget_L := StrMenu(TargetStringMenuQstTxt);
            case SelectedTarget_L of
                1:
                    CopyDistributionLines4Customer(SrcDistributionLine_PT, '');
                2:
                    CopyDistributionLines4Vendors(SrcDistributionLine_PT);
                3:
                    CopyDistributionLines4Contacts(SrcDistributionLine_PT);
                // APO.030 MTH 13.02.24 ...
                4:
                    CopyDistributionLines4RespCenter(SrcDistributionLine_PT);
                // 4: ERROR('');
                // ... APO.030 MTH 13.02.24
                5:
                    Error('');
                6:
                    CopyDistributionLines2Template(SrcDistributionLine_PT);
                // APO.020 SER 07.08.23 ...
                7:
                    CopyDistributionLines2LocationCode(SrcDistributionLine_PT);
                8:
                    CopyDistributionLines2InternalCompany(SrcDistributionLine_PT);
            // ... APO.020 SER 07.08.23
            end;
        end;
    end;

    local procedure CopyDistributionLines4Vendors(SrcDistributionLine_PT: Record "Email Distribution Setup")
    var
        Vendor_LT: Record Vendor;
        TargetDistrLine_LT: Record "Email Distribution Setup";
        DialogMgt_LC: Codeunit "Dialog Mgt.";
        VendorList_LP: Page "Vendor List";
        Copy4VendorsDlgTxt: Label 'Copy for Vendors';
    begin
        Clear(VendorList_LP);
        Clear(DialogMgt_LC);
        Vendor_LT.Reset();
        VendorList_LP.LookupMode := true;
        if VendorList_LP.RunModal() = Action::LookupOK then begin
            VendorList_LP.SetSelection(Vendor_LT);
            DialogMgt_LC.OpenDialog(Vendor_LT.Count(), Copy4VendorsDlgTxt);
            if Vendor_LT.FindSet() then begin
                repeat
                    TargetDistrLine_LT.Reset();
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    if FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", Vendor_LT."Language Code", TargetDistrLine_LT."Use for Type"::Vendor, Vendor_LT."No.", TargetDistrLine_LT) then begin
                        TargetDistrLine_LT.Delete(true);
                    end;
                    if FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", '', TargetDistrLine_LT."Use for Type"::Vendor, Vendor_LT."No.", TargetDistrLine_LT) then begin
                        TargetDistrLine_LT.Delete(true);
                    end;
                    TargetDistrLine_LT.Reset();
                    TargetDistrLine_LT.TransferFields(SrcDistributionLine_PT);
                    TargetDistrLine_LT.Validate("Use for Type", TargetDistrLine_LT."Use for Type"::Vendor);
                    TargetDistrLine_LT.Validate("Use for Code", Vendor_LT."No.");
                    TargetDistrLine_LT.Description := CopyStr(Vendor_LT.Name, 1, MaxStrLen(TargetDistrLine_LT.Description));    // APO.033 MTH 25.04.24
                    // TargetDistrLine_LT.VALIDATE("Language Code", '');    // APO.023 MTH 14.09.23
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
                    if Format(SrcDistributionLine_PT."Linked Template for Email Text") <> '' then
                        TargetDistrLine_LT.Validate("Linked Template for Email Text", SrcDistributionLine_PT."Linked Template for Email Text");
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
                    TargetDistrLine_LT.Insert(true);
                    DialogMgt_LC.UpdateDialog(0);
                until Vendor_LT.Next() = 0;
            end;
            DialogMgt_LC.CloseDialog;
        end;
    end;

    procedure CopyDistributionLines4Customer(SrcDistributionLine_PT: Record "Email Distribution Setup"; Copy2CustomerNo_P: Code[20])
    var
        Customer_LT: Record Customer;
        TargetDistrLine_LT: Record "Email Distribution Setup";
        DialogMgt_LC: Codeunit "Dialog Mgt.";
        CustomerList_LP: Page "Customer List";
        Copy4CustomerDlgTxt: Label 'Copy for Customers';
    begin
        Clear(CustomerList_LP);
        Clear(DialogMgt_LC);
        Customer_LT.Reset();
        if Copy2CustomerNo_P <> '' then begin
            Customer_LT.Get(Copy2CustomerNo_P);
            Customer_LT.SetRecFilter();
        end else begin
            CustomerList_LP.LookupMode := true;
            if CustomerList_LP.RunModal() = Action::LookupOK then begin
                CustomerList_LP.SetSelection(Customer_LT);
            end else
                exit;
        end;
        DialogMgt_LC.OpenDialog(Customer_LT.Count(), Copy4CustomerDlgTxt);
        if Customer_LT.FindSet() then begin
            repeat
                TargetDistrLine_LT.Reset();
                SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                if FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", Customer_LT."Language Code", TargetDistrLine_LT."Use for Type"::Customer, Customer_LT."No.", TargetDistrLine_LT) then begin
                    TargetDistrLine_LT.Delete(true);
                end;
                if FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", '', TargetDistrLine_LT."Use for Type"::Customer, Customer_LT."No.", TargetDistrLine_LT) then begin
                    TargetDistrLine_LT.Delete(true);
                end;
                TargetDistrLine_LT.Reset();
                TargetDistrLine_LT.TransferFields(SrcDistributionLine_PT);
                TargetDistrLine_LT.Validate("Use for Type", TargetDistrLine_LT."Use for Type"::Customer);
                TargetDistrLine_LT.Validate("Use for Code", Customer_LT."No.");
                TargetDistrLine_LT.Description := CopyStr(Customer_LT.Name, 1, MaxStrLen(TargetDistrLine_LT.Description));    // APO.033 MTH 25.04.24
                // TargetDistrLine_LT.VALIDATE("Language Code", '');    // APO.023 MTH 14.09.23
                SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
                if Format(SrcDistributionLine_PT."Linked Template for Email Text") <> '' then
                    TargetDistrLine_LT.Validate("Linked Template for Email Text", SrcDistributionLine_PT."Linked Template for Email Text");
                TargetDistrLine_LT."Recipient Address" := '';
                TargetDistrLine_LT."Recipient Mailing Group" := '';
                TargetDistrLine_LT.Insert(true);
                DialogMgt_LC.UpdateDialog(0);
            until Customer_LT.Next() = 0;
        end;
        DialogMgt_LC.CloseDialog;
    end;

    local procedure CopyDistributionLines4Contacts(SrcDistributionLine_PT: Record "Email Distribution Setup")
    var
        Contact_LT: Record Contact;
        TargetDistrLine_LT: Record "Email Distribution Setup";
        DialogMgt_LC: Codeunit "Dialog Mgt.";
        ContactList_LP: Page "Contact List";
        Copy4ContactsDlgTxt: Label 'Copy for Vendors';
    begin
        Clear(ContactList_LP);
        Clear(DialogMgt_LC);
        Contact_LT.Reset();
        ContactList_LP.LookupMode := true;
        if ContactList_LP.RunModal() = Action::LookupOK then begin
            ContactList_LP.SetSelection(Contact_LT);
            DialogMgt_LC.OpenDialog(Contact_LT.Count(), Copy4ContactsDlgTxt);
            if Contact_LT.FindSet() then begin
                repeat
                    TargetDistrLine_LT.Reset();
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    if FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", Contact_LT."Language Code", TargetDistrLine_LT."Use for Type"::Contact, Contact_LT."No.", TargetDistrLine_LT) then begin
                        TargetDistrLine_LT.Delete(true);
                    end;
                    if FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", '', TargetDistrLine_LT."Use for Type"::Contact, Contact_LT."No.", TargetDistrLine_LT) then begin
                        TargetDistrLine_LT.Delete(true);
                    end;
                    TargetDistrLine_LT.Reset();
                    TargetDistrLine_LT.TransferFields(SrcDistributionLine_PT);
                    TargetDistrLine_LT.Validate("Use for Type", TargetDistrLine_LT."Use for Type"::Contact);
                    TargetDistrLine_LT.Validate("Use for Code", Contact_LT."No.");
                    TargetDistrLine_LT.Validate("Language Code", '');
                    TargetDistrLine_LT.Description := CopyStr(Contact_LT.Name, 1, MaxStrLen(TargetDistrLine_LT.Description));    // APO.033 MTH 25.04.24
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
                    if Format(SrcDistributionLine_PT."Linked Template for Email Text") <> '' then
                        TargetDistrLine_LT.Validate("Linked Template for Email Text", SrcDistributionLine_PT."Linked Template for Email Text");
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
                    TargetDistrLine_LT.Insert(true);
                    DialogMgt_LC.UpdateDialog(0);
                until Contact_LT.Next() = 0;
            end;
            DialogMgt_LC.CloseDialog;
        end;
    end;

    local procedure CopyDistributionLines4RespCenter(SrcDistributionLine_PT: Record "Email Distribution Setup")
    var
        Language_LT: Record Language;
        ResponsibilityCenter_LT: Record "Responsibility Center";
        TargetDistrLine_LT: Record "Email Distribution Setup";
        DialogMgt_LC: Codeunit "Dialog Mgt.";
        ResponsibilityCenterList_LP: Page "Responsibility Center List";
        LangCode_L: Code[10];
        Copy4RespCenter_Txt: Label 'Copy for responsibility center';
    begin
        Clear(ResponsibilityCenterList_LP);
        Clear(DialogMgt_LC);
        ResponsibilityCenter_LT.Reset();
        ResponsibilityCenterList_LP.LookupMode := true;
        if ResponsibilityCenterList_LP.RunModal() = Action::LookupOK then begin
            ResponsibilityCenter_LT.SetFilter(Code, ResponsibilityCenterList_LP.GetSelectionFilter());
            DialogMgt_LC.OpenDialog(ResponsibilityCenter_LT.Count(), Copy4RespCenter_Txt);
            if ResponsibilityCenter_LT.FindSet() then begin
                // APO.033 MTH 25.04.24 ...
                Clear(LangCode_L);
                Language_LT.Reset();
                if Page.RunModal(0, Language_LT) = Action::LookupOK then
                    LangCode_L := Language_LT.Code;
                // ... APO.033 MTH 25.04.24
                repeat
                    TargetDistrLine_LT.Reset();
                    // APO.033 MTH 25.04.24 ...
                    if FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", LangCode_L, TargetDistrLine_LT."Use for Type"::"Responsibility Center", ResponsibilityCenter_LT.Code, TargetDistrLine_LT) then begin
                        TargetDistrLine_LT.Delete(true);
                    end;
                    // SrcDistributionLine_PT.CALCFIELDS("Text for E-Mail Body");
                    // IF FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code",'',TargetDistrLine_LT."Use for Type"::"Responsibility Center",ResponsibilityCenter_LT.Code,TargetDistrLine_LT) THEN BEGIN
                    //   TargetDistrLine_LT.DELETE(TRUE);
                    // END;
                    // ... APO.033 MTH 25.04.24
                    TargetDistrLine_LT.Reset();
                    TargetDistrLine_LT.TransferFields(SrcDistributionLine_PT);
                    TargetDistrLine_LT.Validate("Use for Type", TargetDistrLine_LT."Use for Type"::"Responsibility Center");
                    TargetDistrLine_LT.Validate("Use for Code", ResponsibilityCenter_LT.Code);
                    // APO.033 MTH 25.04.24 ...
                    TargetDistrLine_LT.Description := CopyStr(ResponsibilityCenter_LT.Name, 1, MaxStrLen(TargetDistrLine_LT.Description));
                    TargetDistrLine_LT.Validate("Language Code", LangCode_L);
                    // ... APO.033 MTH 25.04.24
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
                    if Format(SrcDistributionLine_PT."Linked Template for Email Text") <> '' then
                        TargetDistrLine_LT.Validate("Linked Template for Email Text", SrcDistributionLine_PT."Linked Template for Email Text");
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
                    TargetDistrLine_LT.Insert(true);
                    DialogMgt_LC.UpdateDialog(0);
                until ResponsibilityCenter_LT.Next() = 0;
            end;
            DialogMgt_LC.CloseDialog;
        end;
    end;

    local procedure CopyDistributionLines2Template(SrcDistributionLine_PT: Record "Email Distribution Setup")
    var
        FilteredKeyValuePairCodes_LT: Record "Key Value Pair Codes";
        KeyValuePairCodes_LT: Record "Key Value Pair Codes";
        TargetDistrLine_LT: Record "Email Distribution Setup";
        KeyValuePairSelection_LP: Page "Key Value Pair Selection";
        Copy4VendorsDlgTxt: Label 'Copy for Vendors';
    begin
        Clear(KeyValuePairSelection_LP);
        FilteredKeyValuePairCodes_LT.Reset();
        FilteredKeyValuePairCodes_LT.SetRange(Type, FilteredKeyValuePairCodes_LT.Type::"Key Value Pair");
        FilteredKeyValuePairCodes_LT.SetRange("Table ID", Database::"Email Distribution Setup");
        FilteredKeyValuePairCodes_LT.SetRange("Field ID", SrcDistributionLine_PT.FieldNo("Use for Code"));
        KeyValuePairSelection_LP.Editable(false);
        KeyValuePairSelection_LP.LookupMode(true);
        KeyValuePairSelection_LP.SetTableView(FilteredKeyValuePairCodes_LT);
        if KeyValuePairSelection_LP.RunModal() = Action::LookupOK then begin
            KeyValuePairSelection_LP.GetRecord(KeyValuePairCodes_LT);
        end else
            exit;
        TargetDistrLine_LT.Reset();
        TargetDistrLine_LT.TransferFields(SrcDistributionLine_PT);
        TargetDistrLine_LT.Validate("Use for Type", TargetDistrLine_LT."Use for Type"::Template);
        TargetDistrLine_LT.Validate("Use for Code", KeyValuePairCodes_LT.Code);
        TargetDistrLine_LT.Validate("Language Code", '');
        SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
        TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
        TargetDistrLine_LT.Insert(true);
    end;

    local procedure CopyDistributionLines2LocationCode(SrcDistributionLine_PT: Record "Email Distribution Setup")
    var
        Language_LT: Record Language;
        Location_LT: Record Location;
        TargetDistrLine_LT: Record "Email Distribution Setup";
        DialogMgt_LC: Codeunit "Dialog Mgt.";
        LocationList_LP: Page "Location List";
        LangCode_L: Code[10];
        Copy4LocationDlgTxt: Label 'Copy for Location Code';
    begin
        Clear(LocationList_LP);
        Clear(DialogMgt_LC);
        Location_LT.Reset();
        LocationList_LP.LookupMode := true;
        if LocationList_LP.RunModal() = Action::LookupOK then begin
            Location_LT.SetFilter(Code, LocationList_LP.GetSelectionFilter());
            DialogMgt_LC.OpenDialog(Location_LT.Count(), Copy4LocationDlgTxt);
            if Location_LT.FindSet() then begin
                // APO.033 MTH 25.04.24 ...
                Clear(LangCode_L);
                Language_LT.Reset();
                if Page.RunModal(0, Language_LT) = Action::LookupOK then
                    LangCode_L := Language_LT.Code;
                // ... APO.033 MTH 25.04.24
                repeat
                    TargetDistrLine_LT.Reset();
                    // APO.033 MTH 25.04.24 ...
                    if FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", LangCode_L, TargetDistrLine_LT."Use for Type"::"Location Code", Location_LT.Code, TargetDistrLine_LT) then begin
                        TargetDistrLine_LT.Delete(true);
                    end;
                    // SrcDistributionLine_PT.CALCFIELDS("Text for E-Mail Body");
                    // IF FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", '', TargetDistrLine_LT."Use for Type"::"Location Code", Location_LT.Code, TargetDistrLine_LT) THEN BEGIN
                    //   TargetDistrLine_LT.DELETE(TRUE);
                    // END;
                    // IF FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", '', TargetDistrLine_LT."Use for Type"::"Location Code", Location_LT.Code, TargetDistrLine_LT) THEN BEGIN
                    //   TargetDistrLine_LT.DELETE(TRUE);
                    // END;
                    // ... APO.033 MTH 25.04.24
                    TargetDistrLine_LT.Reset();
                    TargetDistrLine_LT.TransferFields(SrcDistributionLine_PT);
                    TargetDistrLine_LT.Validate("Use for Type", TargetDistrLine_LT."Use for Type"::"Location Code");
                    TargetDistrLine_LT.Validate("Use for Code", Location_LT.Code);
                    // APO.033 MTH 25.04.24 ...
                    TargetDistrLine_LT.Description := CopyStr(Location_LT.Name, 1, MaxStrLen(TargetDistrLine_LT.Description));
                    TargetDistrLine_LT.Validate("Language Code", LangCode_L);
                    // TargetDistrLine_LT.VALIDATE("Language Code", '');
                    // ... APO.033 MTH 25.04.24
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
                    if Format(SrcDistributionLine_PT."Linked Template for Email Text") <> '' then
                        TargetDistrLine_LT.Validate("Linked Template for Email Text", SrcDistributionLine_PT."Linked Template for Email Text");
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
                    TargetDistrLine_LT.Insert(true);
                    DialogMgt_LC.UpdateDialog(0);
                until Location_LT.Next() = 0;
            end;
            DialogMgt_LC.CloseDialog;
        end;
    end;

    local procedure CopyDistributionLines2InternalCompany(SrcDistributionLine_PT: Record "Email Distribution Setup")
    var
        Language_LT: Record Language;
        CompanyInformation_LT: Record "Company Information";
        TargetDistrLine_LT: Record "Email Distribution Setup";
        DialogMgt_LC: Codeunit "Dialog Mgt.";
        CompanyList_LP: Page "Company List";
        LangCode_L: Code[10];
        Copy4CompInfoDlgTxt: Label 'Copy for Companies';
    begin
        Clear(CompanyList_LP);
        Clear(DialogMgt_LC);
        CompanyInformation_LT.Reset();
        CompanyList_LP.LookupMode := true;
        if CompanyList_LP.RunModal() = Action::LookupOK then begin
            CompanyInformation_LT.SetFilter(Name, CompanyList_LP.GetSelectionFilter()); // use field "Name" from T79 instead from "Primary Key", because function can not work with empty PK-fields
            DialogMgt_LC.OpenDialog(CompanyInformation_LT.Count(), Copy4CompInfoDlgTxt);
            if CompanyInformation_LT.FindSet() then begin
                // APO.033 MTH 25.04.24 ...
                Clear(LangCode_L);
                Language_LT.Reset();
                if Page.RunModal(0, Language_LT) = Action::LookupOK then
                    LangCode_L := Language_LT.Code;
                // ... APO.033 MTH 25.04.24
                repeat
                    TargetDistrLine_LT.Reset();
                    // APO.033 MTH 25.04.24 ...
                    if FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", LangCode_L, TargetDistrLine_LT."Use for Type"::"Internal Company", CompanyInformation_LT."Primary Key", TargetDistrLine_LT) then begin
                        TargetDistrLine_LT.Delete(true);
                    end;
                    // SrcDistributionLine_PT.CALCFIELDS("Text for E-Mail Body");
                    // IF FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", '', TargetDistrLine_LT."Use for Type"::Contact, CompanyInformation_LT."Primary Key", TargetDistrLine_LT) THEN BEGIN
                    //   TargetDistrLine_LT.DELETE(TRUE);
                    // END;
                    // IF FindDistributionLine(SrcDistributionLine_PT."Email Distribution Code", '', TargetDistrLine_LT."Use for Type"::Contact, CompanyInformation_LT."Primary Key", TargetDistrLine_LT) THEN BEGIN
                    //   TargetDistrLine_LT.DELETE(TRUE);
                    // END;
                    // ... APO.033 MTH 25.04.24
                    TargetDistrLine_LT.Reset();
                    TargetDistrLine_LT.TransferFields(SrcDistributionLine_PT);
                    TargetDistrLine_LT.Description := CopyStr(CompanyInformation_LT.Name, 1, MaxStrLen(TargetDistrLine_LT.Description));    // APO.033 MTH 25.04.24
                    TargetDistrLine_LT.Validate("Use for Type", TargetDistrLine_LT."Use for Type"::"Internal Company");
                    TargetDistrLine_LT.Validate("Use for Code", CompanyInformation_LT."Primary Key");
                    // APO.033 MTH 25.04.24 ...
                    TargetDistrLine_LT.Description := CopyStr(CompanyInformation_LT.Name, 1, MaxStrLen(TargetDistrLine_LT.Description));
                    TargetDistrLine_LT.Validate("Language Code", LangCode_L);
                    // TargetDistrLine_LT.VALIDATE("Language Code", '');
                    // ... APO.033 MTH 25.04.24
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
                    if Format(SrcDistributionLine_PT."Linked Template for Email Text") <> '' then
                        TargetDistrLine_LT.Validate("Linked Template for Email Text", SrcDistributionLine_PT."Linked Template for Email Text");
                    SrcDistributionLine_PT.CalcFields("Text for E-Mail Body");
                    TargetDistrLine_LT."Text for E-Mail Body" := SrcDistributionLine_PT."Text for E-Mail Body";
                    TargetDistrLine_LT.Insert(true);
                    DialogMgt_LC.UpdateDialog(0);
                until CompanyInformation_LT.Next() = 0;
            end;
            DialogMgt_LC.CloseDialog;
        end;
    end;

    procedure GetPlaceholderKeyValueText4Help() HelpText_L: Text
    var
        SalesPerson: Record "Salesperson/Purchaser";
        SalesHeader: Record "Sales Header";
        PurchaseHeader: Record "Purchase Header";
        CompanyInfo: Record "Company Information";
        ExtendedTextLine_LT: Record "Extended Text Line";
        ShippingAgent_LT: Record "Shipping Agent";
        ReminderLevel_LT: Record "Reminder Level";
        Contact: Record Contact;
        ResponsibilityCenter_LT: Record "Responsibility Center";
        ShippingAgentServices_LT: Record "Shipping Agent Services";
        TempDistributionEntry: Record "Email Distribution Entry" temporary;
        DetailRecordHelpLblTxt: Label '%1 = Detail Data from Records';
        PlaceholderLblTxt: Label '%1 = Table: %2 Field: %3\';
    begin
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[SalesPerson.Name]', SalesPerson.TableName(), SalesPerson.FieldName(Name));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[SalesPerson.JobTitle]', SalesPerson.TableName(), SalesPerson.FieldName("Job Title"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[SalesPerson.PhoneNo]', SalesPerson.TableName(), SalesPerson.FieldName("Phone No."));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[SalesPerson.E-Mail]', SalesPerson.TableName(), SalesPerson.FieldName("E-Mail"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[CompanyInfo.Name]', CompanyInfo.TableName(), CompanyInfo.FieldName(Name));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[CompanyInfo.Address]', CompanyInfo.TableName(), CompanyInfo.FieldName(Address));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[CompanyInfo.PostCode]', CompanyInfo.TableName(), CompanyInfo.FieldName("Post Code"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[CompanyInfo.City]', CompanyInfo.TableName(), CompanyInfo.FieldName(City));
        // APO.021 SER 08.08.23 ...
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[CompanyInfo.Address2]', CompanyInfo.TableName(), CompanyInfo.FieldName("Address 2"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[CompanyInfo.CountryRegionCode]', CompanyInfo.TableName(), CompanyInfo.FieldName("Country/Region Code"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[CompanyInfo.VATRegistrationNo]', CompanyInfo.TableName(), CompanyInfo.FieldName("VAT Registration No."));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[CompanyInfo.MailSignatureText]', CompanyInfo.TableName(), CompanyInfo.FieldName("E-Mail Signature Text Code"));
        // ... APO.021 SER 08.08.23
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[DocumentNo]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("Document No."));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[DocumentDate]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("Document Date"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[EmailAddress]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("E-Mail"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[FaxNo]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("Fax No."));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[LanguageCode]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("Language Code"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[CustomerNo]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("Customer No."));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[VendorNo]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("Vendor No."));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[ContactNo]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("Contact No."));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[RespCenterCode]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("Resp. Center Code"));
        // APO.021 SER 08.08.23 ...
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[RespCenter.Name]', ResponsibilityCenter_LT.TableName(), ResponsibilityCenter_LT.FieldName(Name));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[RespCenter.Address]', ResponsibilityCenter_LT.TableName(), ResponsibilityCenter_LT.FieldName(Address));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[RespCenter.Address2]', ResponsibilityCenter_LT.TableName(), ResponsibilityCenter_LT.FieldName("Address 2"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[RespCenter.City]', ResponsibilityCenter_LT.TableName(), ResponsibilityCenter_LT.FieldName(City));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[RespCenter.PostCode]', ResponsibilityCenter_LT.TableName(), ResponsibilityCenter_LT.FieldName("Post Code"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[RespCenter.PhoneNo]', ResponsibilityCenter_LT.TableName(), ResponsibilityCenter_LT.FieldName("Phone No."));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[RespCenter.E-Mail]', ResponsibilityCenter_LT.TableName(), ResponsibilityCenter_LT.FieldName("E-Mail"));
        // APO.029 JUN 17.01.24 ...
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[RespCenter.MailSignatureText]', ResponsibilityCenter_LT.TableName(), ResponsibilityCenter_LT.FieldName("E-Mail Signature Text Code"));
        // ... APO.029 JUN 17.01.24
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[ItemNo]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("Item No."));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[TransferCode]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("Transfer Code"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[ToName]', TempDistributionEntry.TableName(), TempDistributionEntry.FieldName("To Name"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[DueDate]', PurchaseHeader.TableName() + '/' + SalesHeader.TableName(), '24');
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[PmtDiscountPercent]', PurchaseHeader.TableName() + '/' + SalesHeader.TableName(), '25');
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[PmtDiscountDate]', PurchaseHeader.TableName() + '/' + SalesHeader.TableName(), '26');
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[AmountInclVat]', PurchaseHeader.TableName() + '/' + SalesHeader.TableName(), '61');
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[Currency]', PurchaseHeader.TableName() + '/' + SalesHeader.TableName(), 'Currency Code');
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[Salutation]', Contact.TableName(), Contact.FieldName("Salutation Code"));
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[ExternalDocumentNo]', PurchaseHeader.TableName() + '/' + SalesHeader.TableName(), 'External Document No.');
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[OrderDate]', PurchaseHeader.TableName() + '/' + SalesHeader.TableName(), 'Order Date');
        HelpText_L += StrSubstNo(DetailRecordHelpLblTxt, '[DetailsFromSendingRecord]');
        // APO.012 JUN 26.10.22 ...
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[ShippingAgent]', ShippingAgent_LT.TableName(), 'Name');
        // ... APO.012 JUN 26.10.22
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[TrackingLink]', ShippingAgentServices_LT.TableName(), 'Tracking Link');
        // APO.039 JUN 03.09.24 ...
        HelpText_L += StrSubstNo(PlaceholderLblTxt, '[ReminderSubjectLine]', ReminderLevel_LT.TableName(), ReminderLevel_LT.FieldName("Email Subject Text Code"));
        // ... APO.039 JUN 03.09.24



    end;

    local procedure PrepareAddtionalAttachments(EmailDistributionSetupLine_PT: Record "Email Distribution Setup"; TempDistributionEntry_PT: Record "Email Distribution Entry" temporary; RecRef_P: RecordRef; var ExtraFilePathList_P: DotNet List_Of_T)
    var
        ElectronicDocumentFormat_LT: Record "Electronic Document Format";
        IssuedReminderHeader_LT: Record "Issued Reminder Header";
        EmailDistributionAttachment_LT: Record "Email Distribution Attachment";
        FileManagement_LC: Codeunit "File Management";
        MailManagement_LC: Codeunit "Mail Management";
        ApoSalesMgt_LC: Codeunit "Apo Sales Mgt.";
        FinalServerFile_L: File;
        ClientFilePath_L: Text[250];
        NewServerFilePath_L: Text[250];
        ServerFilePath_L: Text[250];
        i: Integer;
    begin
        // APO.002 SER 24.08.22 ...
        // APO.037 JUN 31.07.24 ...
        ExtraFilePathList_P.Clear;
        //CLEAR(ExtraFilePaths_P);
        if EmailDistributionSetupLine_PT."Optional Electr. Doc. Format" <> '' then begin
            ElectronicDocumentFormat_LT.SendElectronically(ServerFilePath_L, ClientFilePath_L, RecRef_P, EmailDistributionSetupLine_PT."Optional Electr. Doc. Format");
            ClientFilePath_L := FileManagement_LC.ClientTempFileName(FileManagement_LC.GetExtension(ServerFilePath_L));
            FileManagement_LC.DownloadHandler(ServerFilePath_L, ClientFilePath_L);
            ClientFilePath_L := FileManagement_LC.MoveAndRenameClientFile(ClientFilePath_L, GetDocumentFileName(RecRef_P, TempDistributionEntry_PT, EmailDistributionSetupLine_PT, FileManagement_LC.GetExtension(ClientFilePath_L)), '');
            // APO.037 JUN 31.07.24 ...
            ExtraFilePathList_P.Add(ClientFilePath_L);
            // statt:
            //ExtraFilePaths_P[1] := ClientFilePath_L;
            // ... APO.037 JUN 31.07.24
            i := 2;
        end;
        // APO.038 JUN 23.08.24 ...
        case RecRef_P.Number() of
            Database::"Issued Reminder Header":
                RecRef_P.SetTable(IssuedReminderHeader_LT);
            ApoSalesMgt_LC.PrepareIssuedReminderEMailAttachmentInvoices(IssuedReminderHeader_LT);
            EmailDistributionAttachment_LT.SetRange("Use for Type", EmailDistributionSetupLine_PT."Use for Type"::Customer);
            EmailDistributionAttachment_LT.SetRange("Use for Code", IssuedReminderHeader_LT."Customer No.");
            EmailDistributionAttachment_LT.SetRange("Language Code", EmailDistributionSetupLine_PT."Language Code");
        end;
        else begin
            EmailDistributionAttachment_LT.SetRange("Use for Type", EmailDistributionSetupLine_PT."Use for Type");
            EmailDistributionAttachment_LT.SetRange("Use for Code", EmailDistributionSetupLine_PT."Use for Code");
            EmailDistributionAttachment_LT.SetRange("Language Code", EmailDistributionSetupLine_PT."Language Code");
        end;
    end;
        // ... APO.038 JUN 23.08.24
        EmailDistributionAttachment_LT.SetRange("Email Distribution Code", EmailDistributionSetupLine_PT."Email Distribution Code");
        // APO.038 JUN 27.08.24 ...
        //EmailDistributionAttachment_LT.SETRANGE("Use for Type",EmailDistributionSetupLine_PT."Use for Type");
        //EmailDistributionAttachment_LT.SETRANGE("Use for Code",EmailDistributionSetupLine_PT."Use for Code");
        // ... APO.038 JUN 27.08.24
        EmailDistributionAttachment_LT.SetFilter("Ending Date", '%1|%2..', 0D, Today());
            repeat
                ServerFilePath_L := FileManagement_LC.ServerTempFileName('mail');
                EmailDistributionAttachment_LT."File BLOB".Export(ServerFilePath_L);
                // APO.041 SER 10.09.24 ...
                if CurrentClientType() in [ClientType::Background, ClientType::NAS] then begin
                    Clear(FinalServerFile_L);
                    FinalServerFile_L.Create(EmailDistributionAttachment_LT."Display File Name");
                    NewServerFilePath_L := FinalServerFile_L.Name();
                    FinalServerFile_L.Close();
                    FileManagement_LC.CopyServerFile(ServerFilePath_L, NewServerFilePath_L, true);
                    ExtraFilePathList_P.Add(NewServerFilePath_L);
                    if FileManagement_LC.ServerFileExists(ServerFilePath_L) then
                        FileManagement_LC.DeleteServerFile(ServerFilePath_L);
                end else begin
                    ClientFilePath_L := FileManagement_LC.DownloadTempFile(ServerFilePath_L);
                    ExtraFilePathList_P.Add(FileManagement_LC.MoveAndRenameClientFile(ClientFilePath_L, EmailDistributionAttachment_LT."Display File Name", ''));
                end;
                // ClientFilePath_L := FileManagement_LC.DownloadTempFile(ServerFilePath_L);
                // // APO.037 JUN 31.07.24 ...
                // ExtraFilePathList_P.Add(FileManagement_LC.MoveAndRenameClientFile(ClientFilePath_L,EmailDistributionAttachment_LT."Display File Name",''));
                // // statt:
                // //ExtraFilePaths_P[i] := FileManagement_LC.MoveAndRenameClientFile(ClientFilePath_L,EmailDistributionAttachment_LT."Display File Name",'');
                // // ... APO.037 JUN 31.07.24
                // ... APO.041 SER 10.09.24
                i += 1;
            until EmailDistributionAttachment_LT.Next() = 0;
        end;
        // ... APO.002 SER 24.08.22


    end;

    procedure CopyAddtionalAttachments2EmailItemFields(var TempEmailItem_PT: Record "Email Item" temporary; var ExtraFilePathList_P: DotNet List_Of_T)
    var
        EmailAttachment_LT: Record "Email Attachment";
        FileManagement_LC: Codeunit "File Management";
        FilePath_L: Text;
        i: Integer;
        FileAdded_L: Boolean;
    begin
        // APO.002 SER 24.08.22 ...
        i := 0;
        // statt:
        //FOR i := 1 TO ARRAYLEN(ExtraFilePaths_P) DO BEGIN
            FileAdded_L := false;
            // APO.037 JUN 31.07.24 ...
            if FilePath_L <> '' then begin
                EmailAttachment_LT.Init();
                EmailAttachment_LT."Email Item ID" := TempEmailItem_PT.ID;
                i += 1;
                EmailAttachment_LT.Number := i;
                // APO.041 SER 10.09.24 ...
                if CurrentClientType() in [ClientType::Background, ClientType::NAS] then begin
                    EmailAttachment_LT."File Path" := FilePath_L;
                end else begin
                    EmailAttachment_LT."File Path" := FileManagement_LC.UploadFileSilentToServerPath(FilePath_L, '');
                end;
                // ... APO.041 SER 10.09.24
                EmailAttachment_LT.Name := FileManagement_LC.GetFileName(FilePath_L);
                EmailAttachment_LT.Insert();
                FileAdded_L := true;
                /*
            // statt:
            //IF ExtraFilePaths_P[i] <> '' THEN BEGIN
            // ... APO.037 JUN 31.07.24
                IF (TempEmailItem_PT."Attachment File Path 2" = '') AND (TempEmailItem_PT."Attachment Name 2" = '') AND (NOT FileAdded_L) THEN BEGIN
                    // APO.037 JUN 31.07.24 ...
                    TempEmailItem_PT."Attachment File Path 2" := FileManagement_LC.UploadFileSilentToServerPath(FilePath_L,'');
                    TempEmailItem_PT."Attachment Name 2" := FileManagement_LC.GetFileName(FilePath_L);
                    // statt:
                    //TempEmailItem_PT."Attachment File Path 2" := FileManagement_LC.UploadFileSilentToServerPath(ExtraFilePaths_P[i],'');
                    //TempEmailItem_PT."Attachment Name 2" := FileManagement_LC.GetFileName(ExtraFilePaths_P[i]);
                    // ... APO.037 JUN 31.07.24
                    FileAdded_L := TRUE;
                END;
                IF (TempEmailItem_PT."Attachment File Path 3" = '') AND (TempEmailItem_PT."Attachment Name 3" = '') AND (NOT FileAdded_L) THEN BEGIN
                    // APO.037 JUN 31.07.24 ...
                    TempEmailItem_PT."Attachment File Path 3" := FileManagement_LC.UploadFileSilentToServerPath(FilePath_L,'');
                    TempEmailItem_PT."Attachment Name 3" := FileManagement_LC.GetFileName(FilePath_L);
                    // statt:
                    //TempEmailItem_PT."Attachment File Path 3" := FileManagement_LC.UploadFileSilentToServerPath(ExtraFilePaths_P[i],'');
                    //TempEmailItem_PT."Attachment Name 3" := FileManagement_LC.GetFileName(ExtraFilePaths_P[i]);
                    // ... APO.037 JUN 31.07.24
                    FileAdded_L := TRUE;
                END;
                IF (TempEmailItem_PT."Attachment File Path 4" = '') AND (TempEmailItem_PT."Attachment Name 4" = '') AND (NOT FileAdded_L) THEN BEGIN
                    // APO.037 JUN 31.07.24 ...
                    TempEmailItem_PT."Attachment File Path 4" := FileManagement_LC.UploadFileSilentToServerPath(FilePath_L,'');
                    TempEmailItem_PT."Attachment Name 4" := FileManagement_LC.GetFileName(FilePath_L);
                    // statt:
                    //TempEmailItem_PT."Attachment File Path 4" := FileManagement_LC.UploadFileSilentToServerPath(ExtraFilePaths_P[i],'');
                    //TempEmailItem_PT."Attachment Name 4" := FileManagement_LC.GetFileName(ExtraFilePaths_P[i]);
                    // ... APO.037 JUN 31.07.24
                    FileAdded_L := TRUE;
                END;
                IF (TempEmailItem_PT."Attachment File Path 5" = '') AND (TempEmailItem_PT."Attachment Name 5" = '') AND (NOT FileAdded_L) THEN BEGIN
                    // APO.037 JUN 31.07.24 ...
                    TempEmailItem_PT."Attachment File Path 5" := FileManagement_LC.UploadFileSilentToServerPath(FilePath_L,'');
                    TempEmailItem_PT."Attachment Name 5" := FileManagement_LC.GetFileName(FilePath_L);
                    // statt:
                    //TempEmailItem_PT."Attachment File Path 5" := FileManagement_LC.UploadFileSilentToServerPath(ExtraFilePaths_P[i],'');
                    //TempEmailItem_PT."Attachment Name 5" := FileManagement_LC.GetFileName(ExtraFilePaths_P[i]);
                    // ... APO.037 JUN 31.07.24
                    FileAdded_L := TRUE;
                END;
                IF (TempEmailItem_PT."Attachment File Path 6" = '') AND (TempEmailItem_PT."Attachment Name 6" = '') AND (NOT FileAdded_L) THEN BEGIN
                    // APO.037 JUN 31.07.24 ...
                    TempEmailItem_PT."Attachment File Path 6" := FileManagement_LC.UploadFileSilentToServerPath(FilePath_L,'');
                    TempEmailItem_PT."Attachment Name 6" := FileManagement_LC.GetFileName(FilePath_L);
                    // statt:
                    //TempEmailItem_PT."Attachment File Path 6" := FileManagement_LC.UploadFileSilentToServerPath(ExtraFilePaths_P[i],'');
                    //TempEmailItem_PT."Attachment Name 6" := FileManagement_LC.GetFileName(ExtraFilePaths_P[i]);
                    // ... APO.037 JUN 31.07.24
                    FileAdded_L := TRUE;
                END;
                IF (TempEmailItem_PT."Attachment File Path 7" = '') AND (TempEmailItem_PT."Attachment Name 7" = '') AND (NOT FileAdded_L) THEN BEGIN
                    // APO.037 JUN 31.07.24 ...
                    TempEmailItem_PT."Attachment File Path 7" := FileManagement_LC.UploadFileSilentToServerPath(FilePath_L,'');
                    TempEmailItem_PT."Attachment Name 7" := FileManagement_LC.GetFileName(FilePath_L);
                    // statt:
                    //TempEmailItem_PT."Attachment File Path 7" := FileManagement_LC.UploadFileSilentToServerPath(ExtraFilePaths_P[i],'');
                    //TempEmailItem_PT."Attachment Name 7" := FileManagement_LC.GetFileName(ExtraFilePaths_P[i]);
                    // ... APO.037 JUN 31.07.24
                    FileAdded_L := TRUE;
                END;
                // APO.037 JUN 31.07.24 ...
                */
                // ... APO.037 JUN 31.07.24
            end;
        end;
        // ... APO.002 SER 24.08.22
    end;

    local procedure AddMailAddressesFromMailingGroup(MailGroupCode_P: Code[10]; MailAddresses_P: Text) NewMailAddresses_P: Text[250]
    var
        NameValueBuffer_LT: Record "Name/Value Buffer" temporary;
        Contact_LT: Record Contact;
        MailingGroup_LT: Record "Mailing Group";
        ContactMailingGroup_LT: Record "Contact Mailing Group";
    begin
        NameValueBuffer_LT.Reset();
        NameValueBuffer_LT.DeleteAll();
        NameValueBuffer_LT.AddNewEntry(MailAddresses_P, MailAddresses_P);
        if MailingGroup_LT.Get(MailGroupCode_P) then begin
            MailingGroup_LT.CalcFields("No. of Contacts");
            if MailingGroup_LT."No. of Contacts" > 0 then begin
                ContactMailingGroup_LT.Reset();
                ContactMailingGroup_LT.SetCurrentKey("Mailing Group Code");
                ContactMailingGroup_LT.SetRange("Mailing Group Code", MailingGroup_LT.Code);
                ContactMailingGroup_LT.FindSet();
                repeat
                    Contact_LT.Get(ContactMailingGroup_LT."Contact No.");
                    if Contact_LT."E-Mail" <> '' then begin
                        NameValueBuffer_LT.Reset();
                        NameValueBuffer_LT.SetRange(Name, Contact_LT."E-Mail");
                        if not NameValueBuffer_LT.FindSet() then
                            NameValueBuffer_LT.AddNewEntry(Contact_LT."E-Mail", Contact_LT."E-Mail");
                    end;
                until ContactMailingGroup_LT.Next() = 0;
            end;
        end;
        NameValueBuffer_LT.Reset();
        if NameValueBuffer_LT.FindSet() then begin
            repeat
                if NewMailAddresses_P = '' then
                    NewMailAddresses_P := NameValueBuffer_LT.Name
                else begin
                    if StrLen(StrSubstNo('%1;%2', NewMailAddresses_P, NameValueBuffer_LT.Name)) <= MaxStrLen(NewMailAddresses_P) then
                        NewMailAddresses_P += StrSubstNo(';%1', NameValueBuffer_LT.Name)
                    else
                        exit(NewMailAddresses_P);
                end;
            until NameValueBuffer_LT.Next() = 0;
        end;
    end;

    procedure GetPlannedSendingDateTime(EmailDistributionSetup_PT: Record "Email Distribution Setup"; ReferenceDateTime_P: DateTime) SendingDateTime_L: DateTime
    var
        PossibleSendingTimesTmp_LT: Record "Activity Log" temporary;
        TypeHelper_LC: Codeunit "Type Helper";
        ReferenceDate_L: Date;
        ReferenceDateTime_L: DateTime;
        TargetDateTime_L: DateTime;
        Counter_L: Integer;
        DayOfWeek_L: Integer;
    begin
        EmailDistributionSetup_PT.TestField("Delay Send E-Mail");
        if ReferenceDateTime_P = 0DT then
            ReferenceDateTime_L := CreateDateTime(Today(), Time())
        else
            ReferenceDateTime_L := ReferenceDateTime_P;
        PossibleSendingTimesTmp_LT.Reset();
        PossibleSendingTimesTmp_LT.DeleteAll();
        if EmailDistributionSetup_PT."Send Mail on Monday" then
            CreatePossibleSendingTimes(PossibleSendingTimesTmp_LT, 1, ReferenceDateTime_L,
                  EmailDistributionSetup_PT."Send Mail on Monday Time 1", EmailDistributionSetup_PT."Send Mail on Monday Time 2", EmailDistributionSetup_PT."Send Mail on Monday Time 3", EmailDistributionSetup_PT."Send Mail on Monday Time 4", EmailDistributionSetup_PT."Send Mail on Monday Time 5");
        if EmailDistributionSetup_PT."Send Mail on Tuesday" then
            CreatePossibleSendingTimes(PossibleSendingTimesTmp_LT, 2, ReferenceDateTime_L,
                  EmailDistributionSetup_PT."Send Mail on Tuesday Time 1", EmailDistributionSetup_PT."Send Mail on Tuesday Time 2", EmailDistributionSetup_PT."Send Mail on Tuesday Time 3", EmailDistributionSetup_PT."Send Mail on Tuesday Time 4", EmailDistributionSetup_PT."Send Mail on Tuesday Time 5");
        if EmailDistributionSetup_PT."Send Mail on Wednesday" then
            CreatePossibleSendingTimes(PossibleSendingTimesTmp_LT, 3, ReferenceDateTime_L,
                  EmailDistributionSetup_PT."Send Mail on Wednesday Time 1", EmailDistributionSetup_PT."Send Mail on Wednesday Time 2", EmailDistributionSetup_PT."Send Mail on Wednesday Time 3", EmailDistributionSetup_PT."Send Mail on Wednesday Time 4", EmailDistributionSetup_PT."Send Mail on Wednesday Time 5");
        if EmailDistributionSetup_PT."Send Mail on Thursday" then
            CreatePossibleSendingTimes(PossibleSendingTimesTmp_LT, 4, ReferenceDateTime_L,
                  EmailDistributionSetup_PT."Send Mail on Thursday Time 1", EmailDistributionSetup_PT."Send Mail on Thursday Time 2", EmailDistributionSetup_PT."Send Mail on Thursday Time 3", EmailDistributionSetup_PT."Send Mail on Thursday Time 4", EmailDistributionSetup_PT."Send Mail on Thursday Time 5");
        if EmailDistributionSetup_PT."Send Mail on Friday" then
            CreatePossibleSendingTimes(PossibleSendingTimesTmp_LT, 5, ReferenceDateTime_L,
                  EmailDistributionSetup_PT."Send Mail on Friday Time 1", EmailDistributionSetup_PT."Send Mail on Friday Time 2", EmailDistributionSetup_PT."Send Mail on Friday Time 3", EmailDistributionSetup_PT."Send Mail on Friday Time 4", EmailDistributionSetup_PT."Send Mail on Friday Time 5");
        if EmailDistributionSetup_PT."Send Mail on Saturday" then
            CreatePossibleSendingTimes(PossibleSendingTimesTmp_LT, 6, ReferenceDateTime_L,
                  EmailDistributionSetup_PT."Send Mail on Saturday Time 1", EmailDistributionSetup_PT."Send Mail on Saturday Time 2", EmailDistributionSetup_PT."Send Mail on Saturday Time 3", EmailDistributionSetup_PT."Send Mail on Saturday Time 4", EmailDistributionSetup_PT."Send Mail on Saturday Time 5");
        if EmailDistributionSetup_PT."Send Mail on Sunday" then
            CreatePossibleSendingTimes(PossibleSendingTimesTmp_LT, 7, ReferenceDateTime_L,
                  EmailDistributionSetup_PT."Send Mail on Sunday Time 1", EmailDistributionSetup_PT."Send Mail on Sunday Time 2", EmailDistributionSetup_PT."Send Mail on Sunday Time 3", EmailDistributionSetup_PT."Send Mail on Sunday Time 4", EmailDistributionSetup_PT."Send Mail on Sunday Time 5");
            // APO.038 JUN 23.08.24 ...
        if Format(EmailDistributionSetup_PT."Send Delay in Minutes") <> '' then begin
            ReferenceDate_L := DT2Date(ReferenceDateTime_L);
            DayOfWeek_L := Date2DWY(ReferenceDate_L, 1);
            PossibleSendingTimesTmp_LT.Reset();
            if PossibleSendingTimesTmp_LT.FindLast() then
                Counter_L := PossibleSendingTimesTmp_LT.ID + 1
            else
                Counter_L := 1;
            TargetDateTime_L := TypeHelper_LC.AddMinutesToDateTime(ReferenceDateTime_L, EmailDistributionSetup_PT."Send Delay in Minutes");
            if
                    ((DayOfWeek_L = 1) and (not EmailDistributionSetup_PT."Send Mail on Monday")) or
                    ((DayOfWeek_L = 2) and (not EmailDistributionSetup_PT."Send Mail on Tuesday")) or
                    ((DayOfWeek_L = 3) and (not EmailDistributionSetup_PT."Send Mail on Wednesday")) or
                    ((DayOfWeek_L = 4) and (not EmailDistributionSetup_PT."Send Mail on Thursday")) or
                    ((DayOfWeek_L = 5) and (not EmailDistributionSetup_PT."Send Mail on Friday")) or
                    ((DayOfWeek_L = 6) and (not EmailDistributionSetup_PT."Send Mail on Saturday")) or
                    ((DayOfWeek_L = 7) and (not EmailDistributionSetup_PT."Send Mail on Sunday")) then
                begin
                    PossibleSendingTimesTmp_LT.Init();
                    PossibleSendingTimesTmp_LT.ID := Counter_L;
                    PossibleSendingTimesTmp_LT."Activity Date" := TargetDateTime_L;
                    PossibleSendingTimesTmp_LT.Insert();
                end;
        end;
            // ... APO.038 JUN 23.08.24
        PossibleSendingTimesTmp_LT.Reset();
        PossibleSendingTimesTmp_LT.SetCurrentKey("Activity Date");
        PossibleSendingTimesTmp_LT.SetFilter("Activity Date", '>=%1', ReferenceDateTime_L);
        if PossibleSendingTimesTmp_LT.FindFirst() then
            exit(PossibleSendingTimesTmp_LT."Activity Date");
    end;

    local procedure CreatePossibleSendingTimes(var PossibleSendingDateTimesTmp_PT: Record "Activity Log" temporary; Weekday_P: Integer; ReferenceDateTime_P: DateTime; Time1_P: Time; Time2_P: Time; Time3_P: Time; Time4_P: Time; Time5_P: Time)
    var
        Date4Calc_L: Date;
        ReferenceDate_L: Date;
        Time4Calc_L: array[5] of Time;
        EntryCounter: Integer;
        i: Integer;
        NoOfDays: Integer;
        NoOfExtraDays: Integer;
        StartingWeekDay: Integer;
        Found: Boolean;
        RunOnDate: array[7] of Boolean;
        CalcDateTxt: Label '+%1D';
    begin
        Clear(Time4Calc_L);
        Time4Calc_L[1] := Time1_P;
        Time4Calc_L[2] := Time2_P;
        Time4Calc_L[3] := Time3_P;
        Time4Calc_L[4] := Time4_P;
        Time4Calc_L[5] := Time5_P;
        ReferenceDate_L := DT2Date(ReferenceDateTime_P);
        PossibleSendingDateTimesTmp_PT.Reset();
        if PossibleSendingDateTimesTmp_PT.FindLast() then
            EntryCounter := PossibleSendingDateTimesTmp_PT.ID + 1
        else
            EntryCounter := 1;
        if Weekday_P = Date2DWY(ReferenceDate_L, 1) then begin
            for i := 1 to 5 do begin
                if Time4Calc_L[i] <> 0T then begin
                    PossibleSendingDateTimesTmp_PT.Init();
                    PossibleSendingDateTimesTmp_PT.ID := EntryCounter;
                    PossibleSendingDateTimesTmp_PT."Activity Date" := CreateDateTime(ReferenceDate_L, Time4Calc_L[i]);
                    PossibleSendingDateTimesTmp_PT.Insert(false);
                    EntryCounter += 1;
                end;
            end;
            for i := 1 to 5 do begin
                if Time4Calc_L[i] <> 0T then begin
                    PossibleSendingDateTimesTmp_PT.Init();
                    PossibleSendingDateTimesTmp_PT.ID := EntryCounter;
                    Date4Calc_L := CalcDate(StrSubstNo(CalcDateTxt, '7'), ReferenceDate_L);
                    PossibleSendingDateTimesTmp_PT."Activity Date" := CreateDateTime(Date4Calc_L, Time4Calc_L[i]);
                    PossibleSendingDateTimesTmp_PT.Insert(false);
                    EntryCounter += 1;
                end;
            end;
        end else begin
            StartingWeekDay := Date2DWY(DT2Date(ReferenceDateTime_P), 1);
            Clear(RunOnDate);
            RunOnDate[1] := Weekday_P = 1;
            RunOnDate[2] := Weekday_P = 2;
            RunOnDate[3] := Weekday_P = 3;
            RunOnDate[4] := Weekday_P = 4;
            RunOnDate[5] := Weekday_P = 5;
            RunOnDate[6] := Weekday_P = 6;
            RunOnDate[7] := Weekday_P = 7;
            while not Found and (NoOfExtraDays < 7) do begin
                NoOfExtraDays := NoOfExtraDays + 1;
                NoOfDays := NoOfDays + 1;
                Found := RunOnDate[(StartingWeekDay - 1 + NoOfDays) mod 7 + 1];
            end;
            for i := 1 to 5 do begin
                if Time4Calc_L[i] <> 0T then begin
                    PossibleSendingDateTimesTmp_PT.Init();
                    PossibleSendingDateTimesTmp_PT.ID := EntryCounter;
                    Date4Calc_L := CalcDate(StrSubstNo(CalcDateTxt, NoOfDays), ReferenceDate_L);
                    PossibleSendingDateTimesTmp_PT."Activity Date" := CreateDateTime(Date4Calc_L, Time4Calc_L[i]);
                    PossibleSendingDateTimesTmp_PT.Insert(false);
                    EntryCounter += 1;
                end;
            end;
        end;
    end;

    local procedure GetPrimaryKeyFieldValuesAsText(RecordAsVariant_P: Variant) PKFieldsWithValueTxt_L: Text
    var
        DataTypeManagement_LC: Codeunit "Data Type Management";
        RecRef_L: RecordRef;
        FieldRef: FieldRef;
        KeyRef: KeyRef;
        KeyFieldIndex: Integer;
    begin
        PKFieldsWithValueTxt_L := '';
        Clear(RecRef_L);
        DataTypeManagement_LC.GetRecordRef(RecordAsVariant_P, RecRef_L);
        KeyRef := RecRef_L.KeyIndex(1);
        for KeyFieldIndex := 1 to KeyRef.FieldCount() do begin
            FieldRef := KeyRef.FieldIndex(KeyFieldIndex);
            PKFieldsWithValueTxt_L += StrSubstNo('%1: %2 ', FieldRef.Caption(), Format(FieldRef.Value()));
        end;
    end;

    local procedure GetCompanyName(RecRef_P: RecordRef) NewCompanyName_L: Text
    var
        CompanyInfo_LT: Record "Company Information";
        DataTypeMgt_LC: Codeunit "Data Type Management";
        FldRef_L: FieldRef;
        InternalCompanyCode_L: Code[20];
        InternalCompanyFieldNameLblTxt: Label 'Internal Company';
    begin
        NewCompanyName_L := CompanyName();
        Clear(CompanyInfo_LT);
        if DataTypeMgt_LC.FindFieldByName(RecRef_P, FldRef_L, InternalCompanyFieldNameLblTxt) then begin
            if Format(FldRef_L.Value()) <> '' then begin
                InternalCompanyCode_L := FldRef_L.Value();
                if CompanyInfo_LT.Get(InternalCompanyCode_L) then
                    NewCompanyName_L := CompanyInfo_LT.Name;
            end else begin
                CompanyInfo_LT.Get();
                NewCompanyName_L := CompanyInfo_LT.Name;
            end;
        end;
    end;

    [Scope('Cloud')]
    procedure DownloadAttachment(TempEmailItem: Record "Email Item" temporary; FieldNoToDownload_P: Integer)
    var
        FileManagement: Codeunit "File Management";
        DataTypeManagement_LC: Codeunit "Data Type Management";
        AttachmentFilePath_L: Text;
        AttachmentName_L: Text;
    begin
        case FieldNoToDownload_P of
            TempEmailItem.FieldNo("Attachment File Path 2"):
                FileManagement.DownloadHandler(TempEmailItem."Attachment File Path 2", SaveFileDialogTitleMsg, '', SaveFileDialogFilterMsg, TempEmailItem."Attachment Name 2");
                TempEmailItem.FieldNo("Attachment File Path 3"):
                FileManagement.DownloadHandler(TempEmailItem."Attachment File Path 3", SaveFileDialogTitleMsg, '', SaveFileDialogFilterMsg, TempEmailItem."Attachment Name 3");
                TempEmailItem.FieldNo("Attachment File Path 4"):
                FileManagement.DownloadHandler(TempEmailItem."Attachment File Path 4", SaveFileDialogTitleMsg, '', SaveFileDialogFilterMsg, TempEmailItem."Attachment Name 4");
                TempEmailItem.FieldNo("Attachment File Path 5"):
                FileManagement.DownloadHandler(TempEmailItem."Attachment File Path 5", SaveFileDialogTitleMsg, '', SaveFileDialogFilterMsg, TempEmailItem."Attachment Name 5");
                TempEmailItem.FieldNo("Attachment File Path 6"):
                FileManagement.DownloadHandler(TempEmailItem."Attachment File Path 6", SaveFileDialogTitleMsg, '', SaveFileDialogFilterMsg, TempEmailItem."Attachment Name 6");
                TempEmailItem.FieldNo("Attachment File Path 7"):
                FileManagement.DownloadHandler(TempEmailItem."Attachment File Path 7", SaveFileDialogTitleMsg, '', SaveFileDialogFilterMsg, TempEmailItem."Attachment Name 7");
        end;
    end;

    [Scope('Cloud')]
    procedure InsertTempAttachment(var EmailAttachmentTmp_PT: Record "Email Attachment" temporary; EmailItemId_P: Guid; NewNumber_P: Integer; FilePath_P: Text; NewName_P: Text)
    var
        FileManagement_LC: Codeunit "File Management";
        ClientFileName_L: Text;
        ClientFilePath_L: Text;
        FileType_L: Text;
    begin
        EmailAttachmentTmp_PT.Init();
        EmailAttachmentTmp_PT."Email Item ID" := EmailItemId_P;
        EmailAttachmentTmp_PT.Number := NewNumber_P;
        EmailAttachmentTmp_PT.Validate(Number, NewNumber_P);
        ClientFilePath_L := FilePath_P;
        ClientFileName_L := NewName_P;
        if StrLen(NewName_P) > MaxStrLen(EmailAttachmentTmp_PT.Name) then
            ClientFileName_L := StrSubstNo('%1.%2', CopyStr(NewName_P, 1, 45), FileManagement_LC.GetExtension(NewName_P));
        FileType_L := FileManagement_LC.GetExtension(NewName_P);
        EmailAttachmentTmp_PT.Validate("File Path", FileManagement_LC.UploadFileSilentToServerPath(ClientFilePath_L, ''));
        EmailAttachmentTmp_PT.Validate(Name, ClientFileName_L);
        EmailAttachmentTmp_PT.Insert(false);
    end;

    [TryFunction]
    procedure TryGetDistributionLine(DistributionCode: Code[10]; var TempDistributionEntry: Record "Email Distribution Entry" temporary; var EmailDistributionSetupLine: Record "Email Distribution Setup")
    begin
        GetDistributionLine(DistributionCode, TempDistributionEntry, EmailDistributionSetupLine);
    end;

    [TryFunction]
    procedure TryFillTempEmailItem(RecRef_P: RecordRef; ReportId_P: Integer; EmailDistributionSetupLine_PT: Record "Email Distribution Setup"; var TempDistributionEntry_PT: Record "Email Distribution Entry" temporary; var TempEmailItem_PT: Record "Email Item" temporary; RequestPageParameters_P: Text)
    begin
        FillTempEmailItem(RecRef_P, ReportId_P, EmailDistributionSetupLine_PT, TempDistributionEntry_PT, TempEmailItem_PT, RequestPageParameters_P);
    end;

    local procedure PreCheckForFixedEMailRecipient(EmailDistributionSetupLine_PT: Record "Email Distribution Setup"; var TempDistributionEntry_PT: Record "Email Distribution Entry" temporary)
    var
        MailReceiptsFallbackFromDistrHdrMgt_L: Text;
    begin
        MailReceiptsFallbackFromDistrHdrMgt_L := TempDistributionEntry_PT."E-Mail";
        TempDistributionEntry_PT."E-Mail" := '';
        if (EmailDistributionSetupLine_PT."Recipient Type" = EmailDistributionSetupLine_PT."Recipient Type"::Fixed) and (TempDistributionEntry_PT."E-Mail" = '') then begin
            TempDistributionEntry_PT."E-Mail" := '';
            if EmailDistributionSetupLine_PT."Recipient Address" <> '' then begin
                TempDistributionEntry_PT."E-Mail" := EmailDistributionSetupLine_PT."Recipient Address";
            end;
        end;
        if TempDistributionEntry_PT."E-Mail" = '' then
            TempDistributionEntry_PT."E-Mail" := MailReceiptsFallbackFromDistrHdrMgt_L;
    end;

    procedure "--- Filter Fcts. ---"()
    begin
    end;

    procedure SetSilentDistribution(NewSilentDistribution: Boolean)
    begin
        SilentDistribution := NewSilentDistribution;
    end;

    procedure "--- Distribution Log ---"()
    begin
    end;

    procedure ShowDistributionLog(RecRelatedVariant: Variant)
    var
        EmailDistributionEntry: Record "Email Distribution Entry";
        DataTypeMgt: Codeunit "Data Type Management";
        RecRef: RecordRef;
    begin
        DataTypeMgt.GetRecordRef(RecRelatedVariant, RecRef);
        EmailDistributionEntry.SetRange("Source Record ID", RecRef.RecordId());
        Page.RunModal(0, EmailDistributionEntry);
    end;

    local procedure "--- Table Setup ---"()
    begin
    end;

    procedure EmailDistrRecordCollDispatcher(Rec2WriteInMailDistrCollection_P: Variant; FieldID_P: Integer; OldValueAsVariant_P: Variant; NewValueAsVariant_P: Variant; UseDistributionCode_P: Code[10]; UseDistributionLanguageCode_P: Code[10]; UseDistributionUse4Type_P: Option; UseDistributionUse4Code_P: Code[20])
    var
        TempEmailItem_LT: Record "Email Item" temporary;
        EmailDistrRecCollectionHdr_LT: Record "Email Distr. Record Coll.";
        EmailDistrRecCollectionLine_LT: Record "Email Distr. Record Coll.";
        TempDistributionEntry_LT: Record "Email Distribution Entry" temporary;
        EmailDistributionSetupLine_LT: Record "Email Distribution Setup";
        DataTypeManagement_LC: Codeunit "Data Type Management";
        EmailDistrHeaderMgt_LC: Codeunit "Email Distr. Header Mgt.";
        EmailDistrRecCollectionHdrAsRecRef_L: RecordRef;
        Rec2WriteInMailDistrCollectionAsRecRef_L: RecordRef;
        SendingDateTime_L: DateTime;
    begin
        Clear(Rec2WriteInMailDistrCollectionAsRecRef_L);
        DataTypeManagement_LC.GetRecordRef(Rec2WriteInMailDistrCollection_P, Rec2WriteInMailDistrCollectionAsRecRef_L);
        if not GetEmailDistrRecordCollHdr(Rec2WriteInMailDistrCollectionAsRecRef_L, EmailDistrRecCollectionHdr_LT,
                                                                            StrSubstNo('%1|%2', EmailDistrRecCollectionHdr_LT."Email Distrib. Entry Status"::New, EmailDistrRecCollectionHdr_LT."Email Distrib. Entry Status"::"Sending planned")) then begin
            CreateEmailDistrRecordCollHdr(Rec2WriteInMailDistrCollectionAsRecRef_L, EmailDistrRecCollectionHdr_LT, UseDistributionCode_P);
            CreateEmailDistrRecordCollLine(Rec2WriteInMailDistrCollectionAsRecRef_L, EmailDistrRecCollectionHdr_LT, EmailDistrRecCollectionLine_LT, FieldID_P, OldValueAsVariant_P, NewValueAsVariant_P);
            Clear(EmailDistrRecCollectionHdrAsRecRef_L);
            DataTypeManagement_LC.GetRecordRef(EmailDistrRecCollectionHdr_LT, EmailDistrRecCollectionHdrAsRecRef_L);
            EmailDistrHeaderMgt_LC.GetHeaderEntry(EmailDistrRecCollectionHdrAsRecRef_L, TempDistributionEntry_LT);
            if EmailDistributionSetupLine_LT.Get(UseDistributionCode_P, UseDistributionLanguageCode_P, UseDistributionUse4Type_P, UseDistributionUse4Code_P) then begin
                TempEmailItem_LT.ID := CreateGuid();
                FillTempEmailItem(EmailDistrRecCollectionHdrAsRecRef_L, 0, EmailDistributionSetupLine_LT, TempDistributionEntry_LT, TempEmailItem_LT, '');
                if EmailDistributionSetupLine_LT."Delay Send E-Mail" then begin
                    SendingDateTime_L := GetPlannedSendingDateTime(EmailDistributionSetupLine_LT, 0DT);
                    EmailDistrRecCollectionHdr_LT."Email Distribution Entry No." := LogMailDistribution(TempEmailItem_LT, 0, SendingDateTime_L);
                    EmailDistrRecCollectionHdr_LT."Email Distrib. Entry Status" := EmailDistrRecCollectionHdr_LT."Email Distrib. Entry Status"::New;
                    EmailDistrRecCollectionHdr_LT.Modify();
                end else begin
                    if TempEmailItem_LT.Send(EmailDistributionSetupLine_LT."Hide Mail Dialog" or SilentDistribution) then begin
                        EmailDistrRecCollectionHdr_LT."Email Distribution Entry No." := LogMailDistribution(TempEmailItem_LT, 1, CreateDateTime(Today(), Time()));
                        EmailDistrRecCollectionHdr_LT."Email Distrib. Entry Status" := EmailDistrRecCollectionHdr_LT."Email Distrib. Entry Status"::Sent;
                        EmailDistrRecCollectionHdr_LT.Modify();
                    end;
                end;
                EmailDistrRecCollectionLine_LT."Email Distrib. Entry Status" := EmailDistrRecCollectionHdr_LT."Email Distrib. Entry Status";
                EmailDistrRecCollectionLine_LT."Email Distribution Entry No." := EmailDistrRecCollectionHdr_LT."Email Distribution Entry No.";
                EmailDistrRecCollectionLine_LT.Modify();
            end;
        end else begin
            if not GetEmailDistrRecordCollLine(Rec2WriteInMailDistrCollectionAsRecRef_L, EmailDistrRecCollectionHdr_LT, EmailDistrRecCollectionLine_LT) then
                CreateEmailDistrRecordCollLine(Rec2WriteInMailDistrCollectionAsRecRef_L, EmailDistrRecCollectionHdr_LT, EmailDistrRecCollectionLine_LT, FieldID_P, OldValueAsVariant_P, NewValueAsVariant_P);
        end;
    end;

    procedure GetEmailDistrRecordCollHdr(Rec2WriteInMailDistrCollection_P: Variant; var EmailDistrRecordCollHdr_PT: Record "Email Distr. Record Coll."; SentStatusFilterAsText_P: Text) EmailDistrRecordCollHdrFound_L: Boolean
    var
        DataTypeManagement_LC: Codeunit "Data Type Management";
        RecRef_L: RecordRef;
    begin
        DataTypeManagement_LC.GetRecordRef(Rec2WriteInMailDistrCollection_P, RecRef_L);
        EmailDistrRecordCollHdr_PT.Reset();
        EmailDistrRecordCollHdr_PT.SetRange(Type, EmailDistrRecordCollHdr_PT.Type::"Collection Header");
        EmailDistrRecordCollHdr_PT.SetRange("Table ID", RecRef_L.Number());
        if SentStatusFilterAsText_P <> '' then
            EmailDistrRecordCollHdr_PT.SetFilter("Email Distrib. Entry Status", SentStatusFilterAsText_P);
        exit(EmailDistrRecordCollHdr_PT.FindFirst());
    end;

    procedure CreateEmailDistrRecordCollHdr(SearchRecordAsVariant_P: Variant; var EmailDistrRecordCollHdr_PT: Record "Email Distr. Record Coll."; EmailDistributionCode_P: Code[10])
    var
        DataTypeManagement_LC: Codeunit "Data Type Management";
        RecRef_L: RecordRef;
    begin
        DataTypeManagement_LC.GetRecordRef(SearchRecordAsVariant_P, RecRef_L);
        EmailDistrRecordCollHdr_PT.Init();
        EmailDistrRecordCollHdr_PT.Validate(Type, EmailDistrRecordCollHdr_PT.Type::"Collection Header");
        EmailDistrRecordCollHdr_PT."Line No." := 0;
        EmailDistrRecordCollHdr_PT."Table ID" := RecRef_L.Number();
        EmailDistrRecordCollHdr_PT."Email Distribution Code" := EmailDistributionCode_P;
        EmailDistrRecordCollHdr_PT."Email Distrib. Entry Status" := EmailDistrRecordCollHdr_PT."Email Distrib. Entry Status";
        EmailDistrRecordCollHdr_PT."Creation Date/Time" := CurrentDateTime();
        EmailDistrRecordCollHdr_PT."Created by" := UserId();
        EmailDistrRecordCollHdr_PT.Insert(true);
    end;

    procedure GetEmailDistrRecordCollLine(SearchRecordAsVariant_P: Variant; EmailDistrRecordCollHdr_PT: Record "Email Distr. Record Coll."; var EmailDistrRecordCollLine_PT: Record "Email Distr. Record Coll.") EmailDistrRecordCollLineFound_L: Boolean
    var
        DataTypeManagement_LC: Codeunit "Data Type Management";
        RecRef_L: RecordRef;
    begin
        DataTypeManagement_LC.GetRecordRef(SearchRecordAsVariant_P, RecRef_L);
        EmailDistrRecordCollLine_PT.Reset();
        EmailDistrRecordCollLine_PT.SetRange(Type, EmailDistrRecordCollHdr_PT.Type::"Collection Line");
        EmailDistrRecordCollLine_PT.SetRange("No.", EmailDistrRecordCollHdr_PT."No.");
        EmailDistrRecordCollLine_PT.SetRange("Table ID", RecRef_L.Number());
        EmailDistrRecordCollLine_PT.SetRange("Record ID", RecRef_L.RecordId());
        exit(EmailDistrRecordCollLine_PT.FindFirst());
    end;

    procedure CreateEmailDistrRecordCollLine(SearchRecordAsVariant_P: Variant; EmailDistrRecordCollHdr_PT: Record "Email Distr. Record Coll."; var EmailDistrRecordCollLine_PT: Record "Email Distr. Record Coll."; FieldID_P: Integer; OldValueAsVariant_P: Variant; NewValueAsVariant_P: Variant)
    var
        EmailDistrRecordCollLastLine_LT: Record "Email Distr. Record Coll.";
        DataTypeManagement_LC: Codeunit "Data Type Management";
        RecRef_L: RecordRef;
        LineNo_L: Integer;
    begin
        EmailDistrRecordCollLastLine_LT.Reset();
        EmailDistrRecordCollLastLine_LT.SetRange(Type, EmailDistrRecordCollHdr_PT.Type::"Collection Line");
        EmailDistrRecordCollLastLine_LT.SetRange("No.", EmailDistrRecordCollHdr_PT."No.");
        if EmailDistrRecordCollLastLine_LT.FindLast() then
            LineNo_L := EmailDistrRecordCollLastLine_LT."Line No." + 10000
        else
            LineNo_L := 10000;
        DataTypeManagement_LC.GetRecordRef(SearchRecordAsVariant_P, RecRef_L);
        EmailDistrRecordCollLine_PT.Init();
        EmailDistrRecordCollLine_PT.Validate(Type, EmailDistrRecordCollHdr_PT.Type::"Collection Line");
        EmailDistrRecordCollLine_PT."No." := EmailDistrRecordCollHdr_PT."No.";
        EmailDistrRecordCollLine_PT."Line No." := LineNo_L;
        EmailDistrRecordCollLine_PT."Table ID" := RecRef_L.Number();
        EmailDistrRecordCollLine_PT."Record ID" := RecRef_L.RecordId();
        EmailDistrRecordCollLine_PT."Email Distribution Code" := EmailDistrRecordCollHdr_PT."Email Distribution Code";
        EmailDistrRecordCollLine_PT."Email Distribution Entry No." := EmailDistrRecordCollHdr_PT."Email Distribution Entry No.";
        EmailDistrRecordCollLine_PT."Email Distrib. Entry Status" := EmailDistrRecordCollHdr_PT."Email Distrib. Entry Status";
        EmailDistrRecordCollLine_PT."Field ID" := FieldID_P;
        EmailDistrRecordCollLine_PT."Old Value" := Format(OldValueAsVariant_P);
        EmailDistrRecordCollLine_PT."New Value" := Format(NewValueAsVariant_P);
        EmailDistrRecordCollHdr_PT."Creation Date/Time" := CurrentDateTime();
        EmailDistrRecordCollHdr_PT."Created by" := UserId();
        EmailDistrRecordCollLine_PT.Insert(true);
    end;

    procedure OpenTableSetup(var SourceTable: Variant)
    var
        EMailDistribTableSetup: Record "Email Distr. Record Coll.";
        EMailDisitribTableSetups: Page "Email Distrib. Record Collect.";
                                      SourceRecID: RecordId;
                                      TableRecRef: RecordRef;
    begin
        TableRecRef.GetTable(SourceTable);
        SourceRecID := TableRecRef.RecordId();
        EMailDistribTableSetup.Reset();
        EMailDistribTableSetup.SetRange("Table ID", TableRecRef.Number());
        EMailDistribTableSetup.SetRange("Record ID", SourceRecID);
        EMailDisitribTableSetups.SetTableView(EMailDistribTableSetup);
        EMailDisitribTableSetups.RunModal();
    end;

    procedure GetDistributionCodeFromTableSetup(var RecRef: RecordRef; Status: Code[20]; NewStatus: Code[20]): Code[10]
    var
        EMailDistribTableSetup: Record "Email Distr. Record Coll.";
    begin
        EMailDistribTableSetup.Reset();
        EMailDistribTableSetup.SetRange("Record ID", RecRef.RecordId());
        EMailDistribTableSetup.SetRange("Old Value", Status);
        EMailDistribTableSetup.SetRange("New Value", NewStatus);
        if EMailDistribTableSetup.FindFirst() then begin
            exit(EMailDistribTableSetup."Email Distribution Code");
        end;
        EMailDistribTableSetup.SetRange("Old Value", Status);
        EMailDistribTableSetup.SetRange("New Value", '');
        if EMailDistribTableSetup.FindFirst() then begin
            exit(EMailDistribTableSetup."Email Distribution Code");
        end;
        EMailDistribTableSetup.SetRange("Old Value", '');
        EMailDistribTableSetup.SetRange("New Value", NewStatus);
        if EMailDistribTableSetup.FindFirst() then begin
            exit(EMailDistribTableSetup."Email Distribution Code");
        end;
    end;

    local procedure "--- QM ---"()
    begin
    end;

    procedure AddTextString(TextString: Text; NoOfParts: Integer)
    var
        i: Integer;
        NoOfPartsError: Label 'The Number of Arguments has to be between 1 and 10.';
    begin
        if (NoOfParts < 1) or (NoOfParts > 10) then begin
            Error(NoOfPartsError);
        end;
        for i := 1 to NoOfParts do begin
            GTS[i] := SelectStr(i, TextString);
        end;
    end;

    local procedure "--CC--"()
    begin
    end;

    local procedure GetRecFieldsForPurchaseHeader(RecRef: RecordRef; var TextField: array[3] of Text)
    var
        PurchHeader: Record "Purchase Header";
    begin
        if RecRef.Number() = Database::"Purchase Header" then begin
            RecRef.SetTable(PurchHeader);
            TextField[1] := PurchHeader."No.";
            TextField[2] := PurchHeader."Buy-from Vendor No.";
            TextField[2] := PurchHeader."Buy-from Vendor Name";
            TextField[3] := Format(PurchHeader."Order Date");
        end;
    end;

    local procedure "--- Setup Mgt. ---"()
    begin
    end;

    procedure CreateNewEmailDistributionSetup(LinkedRecordAsVariant_P: Variant)
    var
        Customer_LT: Record Customer;
        Vendor_LT: Record Vendor;
        Contact_LT: Record Contact;
        EmailDistributionCode_LT: Record "Email Distribution Code";
        NewEmailDistributionSetup_LT: Record "Email Distribution Setup";
        TemplateEmailDistributionSetup_LT: Record "Email Distribution Setup";
        DataTypeManagement_LC: Codeunit "Data Type Management";
        SelectEmailDistributionSetupTemplateList_LP: Page "Email Distrib. Setup Lookup";
                                                         EmailDistributionCodes_LP: Page "Email Distribution Codes";
                                                         EmailDistributionSetupCard_LP: Page "Email Distribution Setup Card";
                                                         RecRef_L: RecordRef;
                                                         FldRef_L: FieldRef;
    begin
        Clear(EmailDistributionCodes_LP);
        Clear(SelectEmailDistributionSetupTemplateList_LP);
        EmailDistributionCodes_LP.LookupMode(true);
        EmailDistributionCodes_LP.Editable(false);
        if EmailDistributionCodes_LP.RunModal() = Action::LookupOK then begin
            EmailDistributionCodes_LP.GetRecord(EmailDistributionCode_LT);
            TemplateEmailDistributionSetup_LT.FilterGroup(2);
            TemplateEmailDistributionSetup_LT.SetRange("Email Distribution Code", EmailDistributionCode_LT.Code);
            TemplateEmailDistributionSetup_LT.SetRange("Use for Type", TemplateEmailDistributionSetup_LT."Use for Type"::Template);
            SelectEmailDistributionSetupTemplateList_LP.SetTableView(TemplateEmailDistributionSetup_LT);
            SelectEmailDistributionSetupTemplateList_LP.LookupMode(true);
            SelectEmailDistributionSetupTemplateList_LP.Editable(false);
            if SelectEmailDistributionSetupTemplateList_LP.RunModal() <> Action::LookupOK then
                exit
            else
                SelectEmailDistributionSetupTemplateList_LP.GetRecord(TemplateEmailDistributionSetup_LT);
            NewEmailDistributionSetup_LT.Init();
            NewEmailDistributionSetup_LT.TransferFields(TemplateEmailDistributionSetup_LT, false);
            NewEmailDistributionSetup_LT.Validate("Email Distribution Code", EmailDistributionCode_LT.Code);
            DataTypeManagement_LC.GetRecordRef(LinkedRecordAsVariant_P, RecRef_L);
            DataTypeManagement_LC.FindFieldByName(RecRef_L, FldRef_L, 'No.');
            case RecRef_L.Number() of
                Database::Contact: begin
                    Contact_LT.Get(FldRef_L.Value());
                    NewEmailDistributionSetup_LT.Validate("Use for Type", NewEmailDistributionSetup_LT."Use for Type"::Contact);
                    NewEmailDistributionSetup_LT.Validate("Use for Code", Contact_LT."No.");
                    NewEmailDistributionSetup_LT.Description := CopyStr(Contact_LT.Name, 1, MaxStrLen(NewEmailDistributionSetup_LT.Description));   // APO.033 MTH 25.04.24
                end;
                Database::Vendor: begin
                    Vendor_LT.Get(FldRef_L.Value());
                    NewEmailDistributionSetup_LT.Validate("Use for Type", NewEmailDistributionSetup_LT."Use for Type"::Vendor);
                    NewEmailDistributionSetup_LT.Validate("Use for Code", Vendor_LT."No.");
                    NewEmailDistributionSetup_LT.Description := CopyStr(Vendor_LT.Name, 1, MaxStrLen(NewEmailDistributionSetup_LT.Description));   // APO.033 MTH 25.04.24
                end;
                Database::Customer: begin
                    Customer_LT.Get(FldRef_L.Value());
                    NewEmailDistributionSetup_LT.Validate("Use for Type", NewEmailDistributionSetup_LT."Use for Type"::Customer);
                    NewEmailDistributionSetup_LT.Validate("Use for Code", Customer_LT."No.");
                    NewEmailDistributionSetup_LT.Description := CopyStr(Customer_LT.Name, 1, MaxStrLen(NewEmailDistributionSetup_LT.Description));   // APO.033 MTH 25.04.24
                end;
            end;
            TemplateEmailDistributionSetup_LT.CalcFields("Text for E-Mail Body");
            NewEmailDistributionSetup_LT."Text for E-Mail Body" := TemplateEmailDistributionSetup_LT."Text for E-Mail Body";
            if TemplateEmailDistributionSetup_LT.Description = '' then begin
                TemplateEmailDistributionSetup_LT.CalcFields("Email Distribution Description");
                NewEmailDistributionSetup_LT.Description := TemplateEmailDistributionSetup_LT."Email Distribution Description";
            end;
            NewEmailDistributionSetup_LT.Insert();
            Page.Run(Page::"Email Distribution Setup Card", NewEmailDistributionSetup_LT);
        end;
    end;

    procedure CreateNewMailingGroup(LinkedRecordAsVariant_P: Variant; Description4MailGroup_P: Text)
    var
        Customer_LT: Record Customer;
        LinkedVendor_LT: Record Vendor;
        SelectedVendor_LT: Record Vendor;
        Contact_LT: Record Contact;
        FilteredContacts_LT: Record Contact;
        SelectedContacts_LT: Record Contact;
        ContactBusinessRelation_LT: Record "Contact Business Relation";
        MailingGroup_LT: Record "Mailing Group";
        ContactMailingGroup_LT: Record "Contact Mailing Group";
        ApoMasterDataSetup_LT: Record "Apo Master Data Setup";
        NoSeriesManagement_LC: Codeunit NoSeriesManagement;
        DataTypeManagement_LC: Codeunit "Data Type Management";
        ContactList_LP: Page "Contact List";
                            RecRef_L: RecordRef;
                            FldRef_L: array[2] of FieldRef;
                            SelectedContactsCount_L: Integer;
                            MailingGroupBelongs2Type_L: Option " ",Customer,Vendor,Contact;
                            CreateMailGroup4ContactQstTxt: Label 'Do you want create for %1 %2 a Mailing Group with %3 Contacts?';
    begin
        Clear(ContactList_LP);
        FilteredContacts_LT.Reset();
        ContactList_LP.LookupMode(true);
        ContactList_LP.Editable(false);
        Clear(RecRef_L);
        Clear(FldRef_L);
        DataTypeManagement_LC.GetRecordRef(LinkedRecordAsVariant_P, RecRef_L);
        DataTypeManagement_LC.FindFieldByName(RecRef_L, FldRef_L[1], 'No.');
        DataTypeManagement_LC.FindFieldByName(RecRef_L, FldRef_L[2], 'Name');
        case RecRef_L.Number() of
            Database::Contact: begin
                FilteredContacts_LT.SetRange("Company No.", Format(FldRef_L[1].Value()));
                MailingGroupBelongs2Type_L := MailingGroupBelongs2Type_L::Contact;
            end;
            Database::Vendor: begin
                FilteredContacts_LT.SetRange("Company No.", ContactBusinessRelation_LT.GetContactNo(ContactBusinessRelation_LT."Link to Table"::Vendor, Format(FldRef_L[1].Value())));
                SelectedVendor_LT.Get(Format(FldRef_L[1].Value()));
                if (SelectedVendor_LT."No." <> SelectedVendor_LT."Pay-to Vendor No.") and
                    (SelectedVendor_LT."Pay-to Vendor No." <> '') then begin
                    LinkedVendor_LT.Get(SelectedVendor_LT."Pay-to Vendor No.");
                    FilteredContacts_LT.SetFilter("Company No.", '%1|%2', ContactBusinessRelation_LT.GetContactNo(ContactBusinessRelation_LT."Link to Table"::Vendor, SelectedVendor_LT."No."),
                                                                                                                            ContactBusinessRelation_LT.GetContactNo(ContactBusinessRelation_LT."Link to Table"::Vendor, LinkedVendor_LT."No."))
                end;
                MailingGroupBelongs2Type_L := MailingGroupBelongs2Type_L::Vendor;
            end;
            Database::Customer: begin
                FilteredContacts_LT.SetRange("Company No.", ContactBusinessRelation_LT.GetContactNo(ContactBusinessRelation_LT."Link to Table"::Customer, Format(FldRef_L[1].Value())));
                MailingGroupBelongs2Type_L := MailingGroupBelongs2Type_L::Customer;
            end;
        end;
        ContactList_LP.SetTableView(FilteredContacts_LT);
        if ContactList_LP.RunModal() = Action::LookupOK then begin
            SelectedContacts_LT.Reset();
            ContactList_LP.SetSelection(SelectedContacts_LT);
            SelectedContactsCount_L := SelectedContacts_LT.Count();
            if SelectedContactsCount_L > 0 then begin
                if Confirm(StrSubstNo(CreateMailGroup4ContactQstTxt, RecRef_L.Caption(), FldRef_L[1].Value(), SelectedContactsCount_L)) then begin
                    ApoMasterDataSetup_LT.Get();
                    MailingGroup_LT.Init();
                    MailingGroup_LT.Validate(Code, NoSeriesManagement_LC.GetNextNo(ApoMasterDataSetup_LT."Mailing Group Nos.", 0D, true));
                    if Description4MailGroup_P = '' then
                        MailingGroup_LT.Validate(Description, Format(FldRef_L[2].Value()))
                    else
                        MailingGroup_LT.Validate(Description, CopyStr(Description4MailGroup_P, 1, MaxStrLen(MailingGroup_LT.Description)));
                    MailingGroup_LT.Validate("Belongs to Type", MailingGroupBelongs2Type_L);
                    MailingGroup_LT.Validate("Belongs to No.", Format(FldRef_L[1].Value()));
                    MailingGroup_LT.Insert(true);
                    SelectedContacts_LT.FindSet();
                    repeat
                        if not ContactMailingGroup_LT.Get(SelectedContacts_LT."No.", MailingGroup_LT.Code) then begin
                            ContactMailingGroup_LT.Init();
                            ContactMailingGroup_LT.Validate("Contact No.", SelectedContacts_LT."No.");
                            ContactMailingGroup_LT.Validate("Mailing Group Code", MailingGroup_LT.Code);
                            ContactMailingGroup_LT.Insert();
                        end;
                    until SelectedContacts_LT.Next()  = 0;
                    MailingGroup_LT.SetRecFilter();
                    Page.Run(0, MailingGroup_LT);
                end;
            end;
        end;
    end;

    procedure SetMailTextTemplateLinkInDistrSetup(var EmailDistributionSetup_PT: Record "Email Distribution Setup") LinkedSet_L: Boolean
    var
        TemplateEmailDistributionSetup_LT: Record "Email Distribution Setup";
        SelectEmailDistributionSetupTemplateList_LP: Page "Email Distrib. Setup Lookup";
                                                         UseForTypeErrTxt: Label '%1 can not be %2';
    begin
        if EmailDistributionSetup_PT."Use for Type" = EmailDistributionSetup_PT."Use for Type"::Template then begin
            EmailDistributionSetup_PT.Validate("Use Email Text from Template", true);
            EmailDistributionSetup_PT.Validate("Linked Template for Email Text", EmailDistributionSetup_PT.RecordId());
            exit(true);
        end;
        Clear(SelectEmailDistributionSetupTemplateList_LP);
        TemplateEmailDistributionSetup_LT.FilterGroup(2);
        TemplateEmailDistributionSetup_LT.SetRange("Email Distribution Code", EmailDistributionSetup_PT."Email Distribution Code");
        TemplateEmailDistributionSetup_LT.SetRange("Use for Type", TemplateEmailDistributionSetup_LT."Use for Type"::Template);
        SelectEmailDistributionSetupTemplateList_LP.SetTableView(TemplateEmailDistributionSetup_LT);
        SelectEmailDistributionSetupTemplateList_LP.LookupMode(true);
        SelectEmailDistributionSetupTemplateList_LP.Editable(false);
        if SelectEmailDistributionSetupTemplateList_LP.RunModal() <> Action::LookupOK then
            exit(false)
        else
            SelectEmailDistributionSetupTemplateList_LP.GetRecord(TemplateEmailDistributionSetup_LT);
        EmailDistributionSetup_PT.Validate("Use Email Text from Template", true);
        EmailDistributionSetup_PT.Validate("Linked Template for Email Text", TemplateEmailDistributionSetup_LT.RecordId());
        EmailDistributionSetup_PT.Modify();
        exit(true);
    end;
}

