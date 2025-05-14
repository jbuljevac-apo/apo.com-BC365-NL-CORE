codeunit 50000 "Apo Functions Mgt."
{
    trigger OnRun()
    begin
    end;

    procedure FilterRecRef(var RecRef: RecordRef; TableFilter: Text; KeyIndex: Integer)
    var
        Fld: Record "Field";
        FldRef: FieldRef;
        FieldCaption2: Text;
        FieldFilter: Text;
        TableName2: Text;
        Pos: Integer;
    begin
        Pos := 1;
        if GetValue(TableFilter, Pos, ':', TableName2) then begin
            Pos += 1;
            while GetValue(TableFilter, Pos, '=', FieldCaption2) and
                        GetValue(TableFilter, Pos, ',', FieldFilter)
            do begin
                Fld.SetRange(TableNo, RecRef.Number());
                Fld.SetRange("Field Caption", FieldCaption2);
                Fld.SetRange(Enabled, true);
                if Fld.FindFirst() then begin
                    FldRef := RecRef.Field(Fld."No.");
                    FldRef.SetFilter(FieldFilter);
                end;
            end;
        end;
        if KeyIndex <> 0 then
            RecRef.CurrentKeyIndex := KeyIndex;
    end;

    local procedure GetValue(var Value: Text; var Position: Integer; StopChar: Char; var ExitValue: Text): Boolean
    var
        Char: Char;
        Advanced: Boolean;
        FirstCharacter: Boolean;
        Stop: Boolean;
    begin
        ExitValue := '';
        FirstCharacter := true;
        if Position > StrLen(Value) then
            exit(false);
        while not Stop do begin
            Char := Value[Position];
            case true of
                Char = '"':
                    if FirstCharacter then
                        Advanced := true
                    else
                        if Advanced then
                            if Value[Position + 1] = '"' then begin
                                ExitValue += Format(Char);
                                Position += 1;
                            end else begin
                                Stop := true;
                                Position += 1;
                            end
                        else
                            ExitValue += Format(Char);
                Char = StopChar:
                    if Advanced then
                        ExitValue += Format(Char)
                    else
                        Stop := true;
                else
                    ExitValue += Format(Char);
            end;
            FirstCharacter := false;
            Position += 1;
            if Position > StrLen(Value) then
                Stop := true;
        end;
        ExitValue := DelChr(ExitValue, '<>');
        exit(true);
    end;

    var
        UserSetup_GT: Record "User Setup";

    local procedure GetUserSetup()
    begin
        UserSetup_GT.Get(UserId());
    end;

    procedure CheckFiCoUser(ShowError_P: Boolean): Boolean
    begin
    end;


    local procedure "###Kontakt#####"()
    begin
    end;

    procedure CreatePersonContactFromCompany(var CompanyContact_PT: Record Contact; var PersonContact_PT: Record Contact)
    var
        TempPersonContact_LT: Record Contact temporary;
    begin
        Clear(PersonContact_PT);
        PersonContact_PT.Reset();
        TempPersonContact_LT.DeleteAll();
        if (CompanyContact_PT.Type = CompanyContact_PT.Type::Person) then begin
            if not CompanyContact_PT.Get(CompanyContact_PT."Company No.") then begin
                exit;
            end;
        end;
        TempPersonContact_LT.Init();
        TempPersonContact_LT.Copy(CompanyContact_PT);
        TempPersonContact_LT."No." := 'xxxx';
        TempPersonContact_LT."First Name" := '';
        TempPersonContact_LT."Middle Name" := '';
        TempPersonContact_LT.Surname := '';
        TempPersonContact_LT."Job Title" := '';
        TempPersonContact_LT.Initials := '';
        TempPersonContact_LT."Salutation Code" := '';
        TempPersonContact_LT.Type := TempPersonContact_LT.Type::Person;
        TempPersonContact_LT."Phone No." := '';
        TempPersonContact_LT."E-Mail" := '';
        TempPersonContact_LT."Organizational Level Code" := '';
        TempPersonContact_LT.Insert(false);
        if Page.RunModal(Page::"Name Details", TempPersonContact_LT) = Action::LookupOK then begin
            PersonContact_PT.Init();
            PersonContact_PT.TransferFields(TempPersonContact_LT);
            PersonContact_PT."No." := '';
            PersonContact_PT.Insert(true);
        end;
    end;

    procedure GetContactNo4OrgLevelCode(SrcType_P: Option Customer,Vendor; SrcNo_P: Code[10]; OrgLevelCode_P: Code[10]) ContNo_P: Code[10]
    var
        Contact_LT: Record Contact;
        ContBusRelation_LT: Record "Contact Business Relation";
        ContNo4SourceType_L: Code[10];
    begin
        Contact_LT.Reset();
        ContNo4SourceType_L := '';
        case SrcType_P of
            SrcType_P::Customer:
                ContNo4SourceType_L := GetContactBusinessRelation(ContBusRelation_LT."Link to Table"::Customer, SrcNo_P);
            SrcType_P::Vendor:
                ContNo4SourceType_L := GetContactBusinessRelation(ContBusRelation_LT."Link to Table"::Vendor, SrcNo_P);
        end;
        if ContNo4SourceType_L <> '' then begin
            Contact_LT.SetRange("Company No.", ContNo4SourceType_L);
            Contact_LT.SetRange("Organizational Level Code", OrgLevelCode_P);
            Contact_LT.SetRange(Type, Contact_LT.Type::Person);
            if Contact_LT.FindFirst() then begin
                ContNo_P := Contact_LT."No.";
            end;
        end;
        exit(ContNo_P);
    end;

    local procedure GetContactBusinessRelation(LinkToTable: Option; AccountNo: Code[20]): Code[20]
    var
        ContBusRelation: Record "Contact Business Relation";
    begin
        ContBusRelation.SetRange("Link to Table", LinkToTable);
        ContBusRelation.SetRange("No.", AccountNo);
        ContBusRelation.FindFirst();
        exit(ContBusRelation."Contact No.");
    end;


    Procedure LookupUserID(VAR UserName: Code[50]): Boolean
    var
        SID: GUID;
    begin
        EXIT(LookupUser(UserName, SID));
    end;


    local Procedure LookupUser(VAR UserName: Code[50]; VAR SID: GUID): Boolean
    var
        User: Record User;
    begin
        User.RESET;
        User.SETCURRENTKEY("User Name");
        User."User Name" := UserName;
        IF User.FIND('=><') THEN;
        IF PAGE.RUNMODAL(PAGE::Users, User) = ACTION::LookupOK THEN BEGIN
            UserName := User."User Name";
            SID := User."User Security ID";
            EXIT(TRUE);
        END;

        EXIT(FALSE);
    end;
}

