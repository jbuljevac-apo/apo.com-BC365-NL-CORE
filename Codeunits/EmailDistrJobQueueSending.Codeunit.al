codeunit 50058 "Email Distr. Job Queue Sending"
{
    // version ApoNL14.01
    TableNo = "Job Queue Entry";

    trigger OnRun()
    var
        EmailDistributionEntryTmp_LT: Record "Email Distribution Entry" temporary;
        ParameterString: Text;
        StringArr: array[10] of Text;
        MaxEntriesToProcess_L: Integer;
    begin
        Rec.TestField("Parameter String");
        ParameterString := Rec."Parameter String";
        Clear(StringArr);
        ReadParameters(ParameterString, StringArr);
        Evaluate(MaxEntriesToProcess_L, StringArr[1]);
        CollectEmailDistribEntryBuffer(EmailDistributionEntryTmp_LT, MaxEntriesToProcess_L, CurrentDateTime(), StringArr[2]);
        SendCollectedEmailDistribEntryBuffer(EmailDistributionEntryTmp_LT);
        TrsfEmailDistribEntryBuffer2RealEmailDistribEntry(EmailDistributionEntryTmp_LT);
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

    local procedure ClearTmpData(var EmailDistributionEntryTmp_PT: Record "Email Distribution Entry" temporary)
    begin
        EmailDistributionEntryTmp_PT.Reset();
        EmailDistributionEntryTmp_PT.DeleteAll();
    end;

    local procedure CollectEmailDistribEntryBuffer(var EmailDistributionEntryTmp_PT: Record "Email Distribution Entry" temporary; MaxEntriesToProcess_P: Integer; MaxPlanedDateTime_P: DateTime; EmailDistributionCodeFilter_P: Text)
    var
        EmailDistributionEntry_LT: Record "Email Distribution Entry";
    begin
        EmailDistributionEntry_LT.Reset();
        EmailDistributionEntry_LT.SetRange("Sending Date", 0D);
        EmailDistributionEntry_LT.SetRange("Sending Time", 0T);
        EmailDistributionEntry_LT.SetFilter("Planned Sending Date/Time", '..%1', MaxPlanedDateTime_P);
        EmailDistributionEntry_LT.SetRange(Status, EmailDistributionEntry_LT.Status::"Sending planned");
        if EmailDistributionCodeFilter_P <> '' then
            EmailDistributionEntry_LT.SetFilter("Email Distribution Code", EmailDistributionCodeFilter_P);
        if EmailDistributionEntry_LT.FindSet() then begin
            repeat
                EmailDistributionEntryTmp_PT := EmailDistributionEntry_LT;
                EmailDistributionEntryTmp_PT.Insert();
            until EmailDistributionEntry_LT.Next() = 0;
        end;
    end;

    local procedure SendCollectedEmailDistribEntryBuffer(var EmailDistributionEntryTmp_PT: Record "Email Distribution Entry" temporary)
    var
        EmailDistrMgt_LC: Codeunit "Email Distr. Mgt.";
        EmptyMailAddresseserrTxt: Label '%1 and/or %2 empty';
    begin
        EmailDistributionEntryTmp_PT.Reset();
        if EmailDistributionEntryTmp_PT.FindSet() then begin
            repeat
                if (EmailDistributionEntryTmp_PT."Sender Address" = '') or (EmailDistributionEntryTmp_PT."Recipient Address" = '') then begin
                    EmailDistributionEntryTmp_PT.Status := EmailDistributionEntryTmp_PT.Status::Error;
                    EmailDistributionEntryTmp_PT."Error Message" := StrSubstNo(EmptyMailAddresseserrTxt, EmailDistributionEntryTmp_PT.FieldCaption("Sender Address"),
                                                                                                                                                                                            EmailDistributionEntryTmp_PT.FieldCaption("Recipient Address"));
                end else begin
                    EmailDistrMgt_LC.SetSilentDistribution(true);
                    if EmailDistrMgt_LC.DistributeRecordMail(EmailDistributionEntryTmp_PT) then begin
                        EmailDistributionEntryTmp_PT.Status := EmailDistributionEntryTmp_PT.Status::Sent;
                        EmailDistrMgt_LC.SetSendingDateTimeOnEmailDistribEntry(EmailDistributionEntryTmp_PT, Today(), Time(), false);
                    end;
                end;
                EmailDistributionEntryTmp_PT.Modify();
            until EmailDistributionEntryTmp_PT.Next() = 0;
        end;
    end;

    local procedure TrsfEmailDistribEntryBuffer2RealEmailDistribEntry(var EmailDistributionEntryTmp_PT: Record "Email Distribution Entry" temporary)
    var
        EmailDistributionEntry_LT: Record "Email Distribution Entry";
    begin
        EmailDistributionEntryTmp_PT.Reset();
        if EmailDistributionEntryTmp_PT.FindSet() then begin
            repeat
                EmailDistributionEntry_LT.LockTable();
                EmailDistributionEntry_LT.Get(EmailDistributionEntryTmp_PT."Entry No.");
                EmailDistributionEntry_LT."Sending Date" := EmailDistributionEntryTmp_PT."Sending Date";
                EmailDistributionEntry_LT."Sending Time" := EmailDistributionEntryTmp_PT."Sending Time";
                EmailDistributionEntry_LT."Return Message" := EmailDistributionEntryTmp_PT."Return Message";
                EmailDistributionEntry_LT."Error Message" := EmailDistributionEntryTmp_PT."Error Message";
                EmailDistributionEntry_LT.Status := EmailDistributionEntryTmp_PT.Status;
                EmailDistributionEntry_LT.Modify(true);
            until EmailDistributionEntryTmp_PT.Next() = 0;
        end;
    end;
}

