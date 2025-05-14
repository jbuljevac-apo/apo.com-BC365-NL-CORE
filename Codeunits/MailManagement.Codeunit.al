codeunit 59520 "Mail Management"
{
    Permissions = tabledata "Email Distribution Entry" = rimd;

    var
        TempEmailAttachment_GT: Record "Email Attachment" temporary;
        ErrorText_G: Text;
        "### Apo Vars ###": Integer;

    local procedure "### Apo Functions ###"()
    begin
    end;

    local procedure CheckAndGetDistributionSMTPSetup()
    var
        EmailDistributionSetupLine: Record "Email Distribution Setup";
    begin
        if TempEmailItem."E-Mail Distribution Code" <> '' then begin
            EmailDistributionSetupLine.Get(TempEmailItem."E-Mail Distribution Code", TempEmailItem."Language Code", TempEmailItem."Use for Type", TempEmailItem."Use for Code");
            if EmailDistributionSetupLine.Sender = EmailDistributionSetupLine.Sender::Fixed then
                SMTPMail.SetCustomSMTPSetupFromDistributionSetup(EmailDistributionSetupLine);
        end;
    end;

    local procedure DoSendWithDistributionCode(): Boolean
    var
        EmailDistributionSetup_LT: Record "Email Distribution Setup";
    begin
        EmailDistributionSetup_LT.Get(TempEmailItem."E-Mail Distribution Code", TempEmailItem."Language Code", TempEmailItem."Use for Type", TempEmailItem."Use for Code");
        if CurrentClientType() = ClientType::Windows then
            HideMailDialog := EmailDistributionSetup_LT."Hide Mail Dialog";
        DoEdit := HideMailDialog;
        if not HideMailDialog then begin
            if RunMailDialog then begin
                Cancelled := false;
            end else begin
                exit(true);
            end;
        end;
        if EmailDistributionSetup_LT."Mail Type" = EmailDistributionSetup_LT."Mail Type"::SMTP then begin
            exit(SendViaSMTP);
        end else begin
            if DoEdit then begin
                if SendMailOnWinClient then begin
                    exit(true);
                end;
            end;
        end;
    end;

    local procedure SendAndLogMailOnWinClientWithDistributionCode(ClientAttachmentFullName_P: Text): Boolean
    var
        Mail_LC: Codeunit Mail;
        EmailDistrMgt_LC: Codeunit "Email Distr. Mgt.";
        AttachmentFileNameArray_L: array[250] of Text;
        ClientAttachmentFullNameArray_L: array[250] of Text;
        ArrayMembers_L: Integer;
        i: Integer;
    begin
        if TempEmailItem."E-Mail Distribution Code" = '' then
            exit(false)
        else begin
            ClientAttachmentFullNameArray_L[1] := ClientAttachmentFullName_P;
            MailSent := Mail_LC.NewMessageAsyncMultipleAttachments(TempEmailItem."Send to", TempEmailItem."Send CC", TempEmailItem."Send BCC", TempEmailItem.Subject, TempEmailItem.GetBodyText, ClientAttachmentFullNameArray_L, not HideMailDialog);
            ArrayMembers_L := CompressArray(ClientAttachmentFullNameArray_L);
            for i := 1 to ArrayMembers_L do
                FileManagement.DeleteClientFile(ClientAttachmentFullNameArray_L[i]);
        end;
        if not MailSent then begin
            ErrorText_G := Mail_LC.GetErrorDesc();
            EmailDistrMgt_LC.LogMailDistribution(TempEmailItem, 1, 0DT);
        end else
            if TempEmailItem."Email Distribution Entry No." = 0 then
                EmailDistrMgt_LC.LogMailDistribution(TempEmailItem, 1, CreateDateTime(Today(), Time()));
        exit(MailSent);
    end;
}

