codeunit 50008 "Apo Job Queue Task Dispatcher"
{
    // version ApoNL14.06
    TableNo = "Job Queue Entry";

    trigger OnRun()
    var
        DurationParameter_L: Duration;
        ParameterString: Text;
        StringArr: array[10] of Text;
        MaxCounter_L: Integer;
        TableID_L: Integer;
    begin
        Rec.TestField("Parameter String");
        ParameterString := Rec."Parameter String";
        Clear(StringArr);
        ReadParameters(ParameterString, StringArr);
        case StringArr[1] of
            // APO.003 SER 15.10.21 ...
            'CheckJobQueue':
                begin
                    Evaluate(DurationParameter_L, StrSubstNo('%1,%2', StringArr[2], StringArr[3]));
                    CheckJobQueueAndSendMail2NAVAdminUser(DurationParameter_L);
                end;
            // ... APO.003 SER 15.10.21
            // APO.010 SER 12.09.23 ...
            'ProcessingDISSequencesAndReactiveItsPossible':
                ProcessingDISSequencesAndReactiveItsPossible;
            // ... APO.010 SER 12.09.23
            // APO.015 SER 12.04.24 ...
            'ImportCDCDocumentsForDocCat':
                ImportCDCDocumentsForDocCat(StringArr[2]);
        // ... APO.015 SER 12.04.24
        end;
    end;

    local procedure ReadParameters(var ParameterString: Text; var pStringArr: array[10] of Text)
    var
        IntText: Text;
        CommaPos: Integer;
        i: Integer;
    begin
        repeat
            i += 1;
            CommaPos := StrPos(ParameterString, ',');
            if CommaPos > 0 then
                IntText := CopyStr(ParameterString, 1, (CommaPos - 1))
            else
                IntText := ParameterString;
            pStringArr[i] := IntText;
            ParameterString := DelStr(ParameterString, 1, CommaPos);
        until (StrLen(ParameterString) = 0) or (CommaPos = 0);
    end;

    local procedure CheckJobQueueAndSendMail2NAVAdminUser(TimeBetweenMails_P: Duration)
    var
        UserSetup_LT: Record "User Setup";
        SMTPMailSetup_LT: Record "SMTP Mail Setup";
        JobQueueEntry_LT: Record "Job Queue Entry";
        EmailItemTmp_LT: Record "Email Item" temporary;
        ApoSetup_LT: Record "Apo Setup";
        EmailDistributionEntry_LT: Record "Email Distribution Entry";
        User_LT: Record User;
        UserTmp_LT: Record User temporary;
        SMTPMail_LC: Codeunit "SMTP Mail";
        ApoFunctionsMgt_LC: Codeunit "Apo Functions Mgt.";
        EmailDistrMgt_LC: Codeunit "Email Distr. Mgt.";
        DateTime_L: DateTime;
        ChkDuration_L: Duration;
        Duration_L: Duration;
        ErrorInfo_L: Text;
        ErrorMailContent_L: Text;
        LastJobQueueError_L: Text;
        SendToText_L: Text;
        SendMailForJobQueueEntry_L: Boolean;
        JobQueueFailedLblTxt: Label 'Error Job Queue Entry';
        JobQueueLblTxt: Label 'Business Central Job Queue';
        LastJobQueueErrorHeadlineLblTxt: Label '<u>Last Error Message:</u>';
        LastJobQueueErrorUnknownLblTxt: Label 'Error unknown';
        NewAutomationErrLblTxt: Label '<b>Error in Automation:</b>';
        NewHostProcessErrLblTxt: Label '<b>Error in Host Processing:</b>';
        NewJobQueueErrLblTxt: Label '<b>Error in Job Queue Entry:</b>';
        NoSMTPSetupLblTxt: Label 'The SMTP Mail has not been set up';
    begin
        ErrorMailContent_L := '';
        SendToText_L := '';
        EmailItemTmp_LT.Reset();
        EmailItemTmp_LT.DeleteAll();
        UserSetup_LT.Reset();
        UserSetup_LT.SetRange("NAV Administrator", true);
        // APO.007 SER 18.01.23 ...
        repeat
            User_LT.SetRange("User Name", UserSetup_LT."User ID");
            if User_LT.FindFirst() then begin
                UserTmp_LT := User_LT;
                UserTmp_LT."Contact Email" := ApoFunctionsMgt_LC.GetUserMailAddressesFromActDirectory(User_LT, false);
                if (UserTmp_LT."Contact Email" = '') and (UserSetup_LT."E-Mail" <> '') then
                    UserTmp_LT."Contact Email" := UserSetup_LT."E-Mail";
                if UserTmp_LT."Contact Email" <> '' then
                    UserTmp_LT.Insert();
            end;
        until UserSetup_LT.Next() = 0;
    end;
        UserTmp_LT.Reset();
        if UserTmp_LT.FindSet() then begin
            repeat
                if SendToText_L = '' then
                    SendToText_L := UserTmp_LT."Contact Email"
                else
                    SendToText_L += ';' + UserTmp_LT."Contact Email";
            until UserTmp_LT.Next() = 0;
        end;
        // SendToText_L := ApoFunctionsMgt_LC.GetUserMailAddressesFromActDirectory(UserTmp_LT, FALSE);
        // UserSetup_LT.SETFILTER("E-Mail",'<>%1','');
        // IF UserSetup_LT.FINDSET THEN BEGIN
        //  REPEAT
        //    IF SendToText_L = '' THEN
        //      SendToText_L := UserSetup_LT."E-Mail"
        //    ELSE
        //      SendToText_L += ';' + UserSetup_LT."E-Mail";
        //  UNTIL UserSetup_LT.NEXT = 0;
        // END;
        // ... APO.006 SER 01.07.22
            Error(NoSMTPSetupLblTxt);
        SMTPMailSetup_LT.GetSetup();
        JobQueueEntry_LT.Reset();
        JobQueueEntry_LT.SetRange(Status, JobQueueEntry_LT.Status::Error);
        if JobQueueEntry_LT.FindSet() then begin
            repeat
                SendMailForJobQueueEntry_L := false;
                Clear(ChkDuration_L);
                EmailDistributionEntry_LT.Reset();
                EmailDistributionEntry_LT.SetRange("Source Record ID", JobQueueEntry_LT.RecordId());
                if EmailDistributionEntry_LT.FindLast() then begin
                    ChkDuration_L := CreateDateTime(Today(), Time()) - CreateDateTime(EmailDistributionEntry_LT."Sending Date", EmailDistributionEntry_LT."Sending Time");
                    if ChkDuration_L >= TimeBetweenMails_P then
                        SendMailForJobQueueEntry_L := true;
                end else
                    SendMailForJobQueueEntry_L := true;
                if SendMailForJobQueueEntry_L then begin
                    JobQueueEntry_LT.CalcFields("Object Caption to Run");
                    EmailItemTmp_LT.Initialize();
                    EmailItemTmp_LT."From Name" := JobQueueLblTxt;
                    EmailItemTmp_LT."From Address" := SMTPMailSetup_LT."User ID";
                    EmailItemTmp_LT."Send to" := SendToText_L;
                    EmailItemTmp_LT.Subject := CopyStr(StrSubstNo('%1 %2 %3 %4', JobQueueEntry_LT.Status, JobQueueEntry_LT."Object Type to Run", JobQueueEntry_LT."Object ID to Run", JobQueueEntry_LT."Object Caption to Run"), 1, MaxStrLen(EmailItemTmp_LT.Subject));
                    EmailItemTmp_LT."E-Mail Source Record ID" := JobQueueEntry_LT.RecordId();
                    EmailItemTmp_LT.Insert();
                    LastJobQueueError_L := JobQueueEntry_LT.GetErrorMessage();
                    if LastJobQueueError_L = '' then
                        LastJobQueueError_L := LastJobQueueErrorUnknownLblTxt;
                    LastJobQueueError_L := LastJobQueueErrorHeadlineLblTxt + '<br>' + LastJobQueueError_L;
                    if ErrorMailContent_L = '' then
                        ErrorMailContent_L := NewJobQueueErrLblTxt + '<br>' +
                                                                    StrSubstNo('%1 %2 %3 %4', JobQueueEntry_LT.Status, JobQueueEntry_LT."Object Type to Run", JobQueueEntry_LT."Object ID to Run", JobQueueEntry_LT.Description)
                                                                    + '<br>' + LastJobQueueError_L + '<br>' + '<br>'
                    else
                        ErrorMailContent_L += NewJobQueueErrLblTxt + '<br>' +
                                                                    StrSubstNo('%1 %2 %3 %4', JobQueueEntry_LT.Status, JobQueueEntry_LT."Object Type to Run", JobQueueEntry_LT."Object ID to Run", JobQueueEntry_LT.Description)
                                                                    + '<br>' + LastJobQueueError_L + '<br>' + '<br>';
                end;
            until JobQueueEntry_LT.Next() = 0;
        end;
        if (SendToText_L <> '') and (ErrorMailContent_L <> '') then begin
            SMTPMail_LC.CreateMessage(JobQueueLblTxt, SMTPMailSetup_LT."User ID", SendToText_L, JobQueueFailedLblTxt, ErrorMailContent_L, true);
            if SMTPMail_LC.TrySend() then begin
                EmailItemTmp_LT.Reset();
                if EmailItemTmp_LT.FindSet() then begin
                    repeat
                        EmailDistrMgt_LC.LogMailDistribution(EmailItemTmp_LT, 1, CreateDateTime(Today(), Time())); // APO.005 SER 16.12.21
                    until EmailItemTmp_LT.Next() = 0;
                end;
            end;
        end;
        // ... APO.003 SER 15.10.21

    end;

    local procedure ProcessingDISSequencesAndReactiveItsPossible()
    var
        JobQueueEntry_LT: Record "Job Queue Entry";
        ApoDISMgt_LC: Codeunit "Apo DIS Mgt.";
        SequenceAdapter_LC: Codeunit "DIS - Sequence Adapter";
        SeqAdapterInJobQueue_Err: Label 'Codeunit %1 DIS - Sequence Adapter set up as job queue entry. Only one job queue entry allowed to procressing the DIS sequences.';
    begin
        JobQueueEntry_LT.Reset();
        JobQueueEntry_LT.SetRange("Object Type to Run", JobQueueEntry_LT."Object Type to Run"::Codeunit);
        JobQueueEntry_LT.SetRange("Object ID to Run", Codeunit::"DIS - Sequence Adapter");
        if not JobQueueEntry_LT.IsEmpty() then
            Error(StrSubstNo(SeqAdapterInJobQueue_Err, Format(Codeunit::"DIS - Sequence Adapter")));
        Clear(SequenceAdapter_LC);
        if SequenceAdapter_LC.Run() then begin
            Clear(ApoDISMgt_LC);
            ApoDISMgt_LC.ReactiveSequenceHeaders;
            ApoDISMgt_LC.ReactiveSequenceLines;
        end else
            Error(GetLastErrorText());
    end;

    local procedure ImportCDCDocumentsForDocCat(DocCat_P: Code[20])
    var
        DocCat_LT: Record "CDC Document Category";
        ShowDocFilesArgTmp_LT: Record "CDC Show Doc. & Files Arg Tmp";
        TempDisplayDocument_LT: Record "CDC Temp. Display Document" temporary;
        FileSysMgt_LC: Codeunit "CDC File System Management";
        ContiniaOnline_LC: Codeunit "CDC Continia Online";
        FilePath_L: Text[1024];
        Files_L: array[10000] of Text[250];
        EntryNo_L: Integer;
        FileCount_L: Integer;
        i: Integer;
    begin
        DocCat_LT.Get(DocCat_P);
        if not ContiniaOnline_LC.IsCompanyActive2(true, true, true) then
            exit;
        if ContiniaOnline_LC.IsCloudActive(false) then
            ContiniaOnline_LC.GetStatusDocuments(DocCat_LT, ShowDocFilesArgTmp_LT.Status::"Files for Import", TempDisplayDocument_LT)
        else begin
            FilePath_L := DocCat_LT.GetCategoryPath(2);
            if FileSysMgt_LC.DirectoryExists(FilePath_L) then begin
                if TempDisplayDocument_LT.FindLast() then
                    EntryNo_L := TempDisplayDocument_LT."Entry No.";
                FileCount_L := FileSysMgt_LC.GetFilesInDir2(FilePath_L, '*.tiff', Files_L) + FileSysMgt_LC.GetFilesInDir2(FilePath_L, '*.cxml', Files_L);
                for i := 1 to FileCount_L do begin
                    EntryNo_L += 1;
                    CreateTempDocRecordHelper4CDCImport(TempDisplayDocument_LT, DocCat_LT, ShowDocFilesArgTmp_LT.Status::"Files for Import", Files_L[i], EntryNo_L);
                end;
            end;
        end;
        Clear(Files_L);
        FileCount_L := FileSysMgt_LC.GetFilesInDir2(DocCat_LT.GetCategoryPath(4), '*.xml', Files_L);
        if TempDisplayDocument_LT.FindLast() then
            EntryNo_L := TempDisplayDocument_LT."Entry No.";
        for i := 1 to FileCount_L do begin
            EntryNo_L += 1;
            CreateTempDocRecordHelper4CDCImport(TempDisplayDocument_LT, DocCat_LT, ShowDocFilesArgTmp_LT.Status::"Files for Import", Files_L[i], EntryNo_L);
        end;
        TempDisplayDocument_LT.Reset();
        if TempDisplayDocument_LT.FindSet() then
            repeat
                TempDisplayDocument_LT.Import();
            until TempDisplayDocument_LT.Next() = 0;
    end;

    local procedure CreateTempDocRecordHelper4CDCImport(var TempDoc_PT: Record "CDC Temp. Display Document" temporary; DocCat_PT: Record "CDC Document Category"; Status_P: Option "Files for OCR","Files for Import","Files with Error","Open Documents","Registered Documents","Rejected Documents","UIC Documents"; Filename_P: Text[1024]; EntryNo_P: Integer)
    var
        DocumentImporter_LC: Codeunit "CDC Document Importer";
        FileSysMgt_LC: Codeunit "CDC File System Management";
        FileInfo_LC: Codeunit "CDC File Information";
        MetaDocFilePath_L: Text[1024];
    begin
        TempDoc_PT.Init();
        TempDoc_PT."Entry No." := EntryNo_P;
        TempDoc_PT."Scanned File" := true;
        TempDoc_PT."Document Category Code" := DocCat_PT.Code;
        TempDoc_PT."File Name" := FileInfo_LC.GetFilenameWithoutExt(Filename_P);
        TempDoc_PT."File Name with Extension" := FileInfo_LC.GetFilename(Filename_P);
        TempDoc_PT."File Path" := FileInfo_LC.GetFilePath(Filename_P);
        TempDoc_PT."Date/Time" := CurrentDateTime();
        TempDoc_PT.Status := Status_P;
        if TempDoc_PT.Status = TempDoc_PT.Status::Import then
            case true of
                FileSysMgt_LC.FileExists(FileInfo_LC.GetFilePath(Filename_P) + '\' + TempDoc_PT."File Name" + '.xml'):
                    MetaDocFilePath_L := FileInfo_LC.GetFilePath(Filename_P) + '\' + TempDoc_PT."File Name" + '.xml';
                FileSysMgt_LC.FileExists(FileInfo_LC.GetFilePath(Filename_P) + '\' + TempDoc_PT."File Name" + '.cxml'):
                    MetaDocFilePath_L := FileInfo_LC.GetFilePath(Filename_P) + '\' + TempDoc_PT."File Name" + '.cxml';
            end
        else
            if TempDoc_PT.Status = TempDoc_PT.Status::Error then
                MetaDocFilePath_L := DocCat_PT.GetCategoryPath(3) + '\' + TempDoc_PT."File Name with Extension" + '.metadata.xml';
        DocumentImporter_LC.GetDocInfo(MetaDocFilePath_L,
            TempDoc_PT."From E-Mail Address", TempDoc_PT."E-Mail Subject", TempDoc_PT."E-Mail Received", TempDoc_PT."OCR Processed");
        if TempDoc_PT."From E-Mail Address" <> '' then begin
            TempDoc_PT.From := TempDoc_PT."From E-Mail Address";
            TempDoc_PT.Type := TempDoc_PT.Type::Email;
        end else begin
            TempDoc_PT.Type := TempDoc_PT.Type::LocalPath;
            TempDoc_PT.From := FileInfo_LC.GetFilePath(Filename_P);
        end;
        TempDoc_PT.Insert();
    end;
}

