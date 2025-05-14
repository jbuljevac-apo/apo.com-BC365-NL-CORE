codeunit 50018 "Daily Maintain Mgt."
{
    // version ApoDE14.65,ApoNL14.04
    TableNo = "Job Queue Entry";

    trigger OnRun()
    var
        ParameterString_L: Text;
        StringArr_L: array[20] of Text;
    begin
        ParameterString_L := Rec."Parameter String";
        Clear(StringArr_L);
        ReadParameters(ParameterString_L, StringArr_L);

        case StringArr_L[1] of
            // APO.010 JUN 06.08.24 ...
            'CreateReminders':
                CreateReminders(StringArr_L[2]);
            // ... APO.010 JUN 06.08.24
            // APO.011 JUN 26.08.24 ...
            'RegisterReminders':
                RegisterReminders(StringArr_L[2]);
            // ... APO.011 JUN 26.08.24
            else
                Error('Funktionsaufruf nicht implementiert');
        end
        // UpdateCustomers;
    end;

    local procedure CreateReminders(ResponsibilityCenterCode_P: Code[10])
    var
        Customer_LT: Record Customer;
        ResponsibilityCenter_LT: Record "Responsibility Center";
        CreateReminders_LR: Report "Create Reminders";
    begin
        ResponsibilityCenter_LT.Reset();
        if ResponsibilityCenterCode_P <> '' then
            ResponsibilityCenter_LT.SetRange(Code, ResponsibilityCenterCode_P);
        ResponsibilityCenter_LT.SetRange("Create Reminders", true);
        if ResponsibilityCenter_LT.FindSet(false, false) then
            repeat
                Customer_LT.Reset();
                Customer_LT.SetRange("Responsibility Center", ResponsibilityCenter_LT.Code);
                CreateReminders_LR.SetTableView(Customer_LT);
                CreateReminders_LR.InitializeRequest(WorkDate(), WorkDate(), true, true, false);
                CreateReminders_LR.Run();
                Commit();
            until ResponsibilityCenter_LT.Next() = 0;
    end;

    local procedure RegisterReminders(ResponsibilityCenterCode_P: Code[10])
    var
        ReminderHeader_LT: Record "Reminder Header";
        ResponsibilityCenter_LT: Record "Responsibility Center";
        ReminderIssue_LC: Codeunit "Reminder-Issue";
        ErrorMsg_L: Text;
        Error_L: Boolean;
    begin
        ResponsibilityCenter_LT.Reset();
        if ResponsibilityCenterCode_P <> '' then
            ResponsibilityCenter_LT.SetRange(Code, ResponsibilityCenterCode_P);
        ResponsibilityCenter_LT.SetRange("Create Reminders", true);
        if ResponsibilityCenter_LT.FindSet(false, false) then
            repeat
                ReminderHeader_LT.Reset();
                ReminderHeader_LT.SetRange("Responsibility Center", ResponsibilityCenter_LT.Code);
                if ReminderHeader_LT.FindSet(false, false) then
                    repeat
                        Clear(ReminderIssue_LC);
                        ReminderIssue_LC.Set(ReminderHeader_LT, true, WorkDate());
                        if not ReminderIssue_LC.Run() then begin
                            Error_L := true;
                            ErrorMsg_L += GetLastErrorText() + '\n';
                        end;
                        Commit();
                    until ReminderHeader_LT.Next() = 0;
            until ResponsibilityCenter_LT.Next() = 0;
        if Error_L then
            Error(ErrorMsg_L);
    end;

    local procedure ReadParameters(var ParameterString: Text; var pStringArr: array[12] of Text)
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
}

