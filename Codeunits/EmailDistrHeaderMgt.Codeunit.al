codeunit 50060 "Email Distr. Header Mgt."
{
    // version ApoNL14.02
    Permissions = tabledata "Email Distribution Entry" = rimd;

    trigger OnRun()
    begin
    end;

    var
        TextAssemblyType: Label 'Assembly Order';
        TextCheckType: Label 'Check';
        TextDeliveryReminderType: Label 'Delivery Reminder';
        TextEmailDistrRecordColl: Label 'Email Distr. Record Coll.';
        TextReminderType: Label 'Reminder';
        TextServiceContractType: Label 'Service Contract - ';
        TextServiceCrMemoType: Label 'Service Credit Memo';
        TextServiceInvoiceType: Label 'Service Invoice';
        TextServiceShptType: Label 'Service Shipment';
        TextServiceType: Label 'Service - ';
        TextShipmentType: Label 'Shipment';
        TextTransferType: Label 'Transfer Order';

    procedure GetHeaderEntry(HeaderDoc: Variant; var TempDistrEntry: Record "Email Distribution Entry" temporary)
    var
        DataTypeMgt: Codeunit "Data Type Management";
        RecRef: RecordRef;
    begin
        if not HeaderDoc.IsRecordRef() then begin
            DataTypeMgt.GetRecordRef(HeaderDoc, RecRef);
            RecRef.Find();
        end else begin
            RecRef := HeaderDoc;
        end;
        TempDistrEntry.DeleteAll();
        TempDistrEntry.Init();
        case RecRef.Number() of
            Database::"Purchase Header",
            Database::"Purch. Cr. Memo Hdr.",
            Database::"Return Shipment Header",
            Database::"Purchase Header Archive":
                GetFromPurchaseDocument(RecRef, TempDistrEntry);
            Database::"Sales Header",
            Database::"Sales Shipment Header",
            Database::"Sales Invoice Header",
            Database::"Sales Cr.Memo Header",
            Database::"Return Receipt Header",
            Database::"Sales Header Archive":
                GetFromSalesDocument(RecRef, TempDistrEntry);
            Database::"Issued Reminder Header",
            Database::"Bank Account Statement",
            Database::"Gen. Journal Line",
            Database::Table5005272,
            Database::Table5005270:
                GetFromFinancialDocument(RecRef, TempDistrEntry);
            Database::"Service Contract Header",
            Database::"Service Header",
            Database::"Service Shipment Header",
            Database::"Service Invoice Header",
            Database::"Service Cr.Memo Header":
                GetFromServiceDocument(RecRef, TempDistrEntry);
            Database::"Transfer Header",
            Database::"Assembly Header",
            Database::"Email Distr. Record Coll.":
                GetFromOtherDocument(RecRef, TempDistrEntry);
        end;

        SetResponsibilityCenter(RecRef, TempDistrEntry);
        SetEmailAndFax(TempDistrEntry);

        TempDistrEntry.Insert();
    end;

    local procedure GetFromPurchaseDocument(RecRef: RecordRef; var TempDistrEntry: Record "Email Distribution Entry" temporary)
    var
        PurchHeader: Record "Purchase Header";
        PurchCrMemoHdr: Record "Purch. Cr. Memo Hdr.";
        PurchHeaderArch: Record "Purchase Header Archive";
        ReturnShptHeader: Record "Return Shipment Header";
    begin
        case RecRef.Number() of
            Database::"Purchase Header":
                begin
                    RecRef.SetTable(PurchHeader);
                    TempDistrEntry."Document Type" := Format(PurchHeader."Document Type");
                    TempDistrEntry."Document No." := PurchHeader."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := ''; // direct shipment will otherwise use customer no.
                    TempDistrEntry."Vendor No." := PurchHeader."Buy-from Vendor No.";
                    TempDistrEntry."To Name" := PurchHeader."Buy-from Vendor Name";
                    TempDistrEntry."Language Code" := PurchHeader."Language Code";
                    TempDistrEntry."Document Date" := PurchHeader."Document Date";
                    TempDistrEntry."Source Record ID" := RecRef.RecordId();
                    TempDistrEntry."Order Address Code" := PurchHeader."Order Address Code";
                    // APO.002 SER 07.08.23 ...
                    TempDistrEntry."Location Code" := PurchHeader."Location Code";
                    TempDistrEntry."Internal Company" := '';
                    // ... APO.002 SER 07.08.23
                end;
            Database::"Purch. Cr. Memo Hdr.":
                begin
                    RecRef.SetTable(PurchCrMemoHdr);
                    TempDistrEntry."Document Type" := Format(PurchHeader."Document Type"::"Credit Memo");
                    TempDistrEntry."Document No." := PurchCrMemoHdr."No.";
                    TempDistrEntry."Contact No." := PurchCrMemoHdr."Pay-to Contact No.";
                    TempDistrEntry."Customer No." := GetFirstSourceTableFieldCode(RecRef, Database::Customer);
                    TempDistrEntry."Vendor No." := PurchCrMemoHdr."Pay-to Vendor No.";
                    TempDistrEntry."Resp. Center Code" := GetFirstSourceTableFieldCode(RecRef, Database::"Responsibility Center");
                    TempDistrEntry."To Name" := PurchCrMemoHdr."Buy-from Vendor Name";
                    TempDistrEntry."Language Code" := PurchCrMemoHdr."Language Code";
                    TempDistrEntry."Document Date" := PurchCrMemoHdr."Document Date";
                    TempDistrEntry."Source Record ID" := RecRef.RecordId();
                    // APO.002 SER 07.08.23 ...
                    TempDistrEntry."Location Code" := PurchCrMemoHdr."Location Code";
                    TempDistrEntry."Internal Company" := '';
                    // ... APO.002 SER 07.08.23
                end;
            Database::"Return Shipment Header":
                begin
                    RecRef.SetTable(ReturnShptHeader);
                    TempDistrEntry."Document Type" := Format(PurchHeader."Document Type"::"Return Order");
                    TempDistrEntry."Document No." := ReturnShptHeader."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := ''; // direct shipment will otherwise use customer no.
                    TempDistrEntry."Vendor No." := ReturnShptHeader."Buy-from Vendor No.";
                    TempDistrEntry."To Name" := ReturnShptHeader."Buy-from Vendor Name";
                    TempDistrEntry."Language Code" := ReturnShptHeader."Language Code";
                    TempDistrEntry."Document Date" := ReturnShptHeader."Document Date";
                    TempDistrEntry."Source Record ID" := RecRef.RecordId();
                    // APO.002 SER 07.08.23 ...
                    TempDistrEntry."Location Code" := ReturnShptHeader."Location Code";
                    TempDistrEntry."Internal Company" := '';
                    // ... APO.002 SER 07.08.23
                end;
            Database::"Purchase Header Archive":
                begin
                    RecRef.SetTable(PurchHeaderArch);
                    TempDistrEntry."Document Type" := Format(PurchHeaderArch."Document Type");
                    TempDistrEntry."Document No." := PurchHeaderArch."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := GetFirstSourceTableFieldCode(RecRef, Database::Customer);
                    TempDistrEntry."Vendor No." := PurchHeaderArch."Pay-to Vendor No.";
                    TempDistrEntry."To Name" := PurchHeaderArch."Pay-to Name";
                    TempDistrEntry."Language Code" := PurchHeaderArch."Language Code";
                    TempDistrEntry."Document Date" := PurchHeaderArch."Document Date";
                    TempDistrEntry."Source Record ID" := RecRef.RecordId();
                    // APO.002 SER 07.08.23 ...
                    TempDistrEntry."Location Code" := PurchHeaderArch."Location Code";
                    TempDistrEntry."Internal Company" := '';
                    // ... APO.002 SER 07.08.23
                end;
        end;
    end;

    local procedure GetFromSalesDocument(RecRef: RecordRef; var TempDistrEntry: Record "Email Distribution Entry" temporary)
    var
        SalesHeader: Record "Sales Header";
        SalesShptHeader: Record "Sales Shipment Header";
        SalesInvoiceHeader: Record "Sales Invoice Header";
        SalesCrMemoHeader: Record "Sales Cr.Memo Header";
        SalesHeaderArch: Record "Sales Header Archive";
        ReturnReceiptHeader: Record "Return Receipt Header";
    begin
        case RecRef.Number() of
            Database::"Sales Header":
                begin
                    RecRef.SetTable(SalesHeader);
                    TempDistrEntry."Document Type" := Format(SalesHeader."Document Type");
                    TempDistrEntry."Document No." := SalesHeader."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := SalesHeader."Bill-to Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := SalesHeader."Bill-to Name";
                    TempDistrEntry."Language Code" := SalesHeader."Language Code";
                    TempDistrEntry."Document Date" := SalesHeader."Document Date";
                    TempDistrEntry."Ship-to Address Code" := SalesHeader."Ship-to Code"; // APO.002 SER 06.12.22
                    // APO.002 SER 07.08.23 ...
                    TempDistrEntry."Location Code" := SalesHeader."Location Code";
                    TempDistrEntry."Internal Company" := '';
                    // ... APO.002 SER 07.08.23
                end;
            Database::"Sales Shipment Header":
                begin
                    RecRef.SetTable(SalesShptHeader);
                    TempDistrEntry."Document Type" := TextShipmentType;
                    TempDistrEntry."Document No." := SalesShptHeader."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := SalesShptHeader."Bill-to Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := SalesShptHeader."Bill-to Name";
                    TempDistrEntry."Language Code" := SalesShptHeader."Language Code";
                    TempDistrEntry."Document Date" := SalesShptHeader."Document Date";
                    // APO.002 SER 07.08.23 ...
                    TempDistrEntry."Location Code" := SalesShptHeader."Location Code";
                    TempDistrEntry."Internal Company" := '';
                    // ... APO.002 SER 07.08.23
                end;
            Database::"Sales Invoice Header":
                begin
                    RecRef.SetTable(SalesInvoiceHeader);
                    TempDistrEntry."Document Type" := Format(SalesHeader."Document Type"::Invoice);
                    TempDistrEntry."Document No." := SalesInvoiceHeader."No.";
                    TempDistrEntry."Contact No." := SalesInvoiceHeader."Bill-to Contact No.";
                    TempDistrEntry."Customer No." := SalesInvoiceHeader."Bill-to Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := SalesInvoiceHeader."Bill-to Name";
                    TempDistrEntry."Language Code" := SalesInvoiceHeader."Language Code";
                    TempDistrEntry."Document Date" := SalesInvoiceHeader."Document Date";
                    // APO.002 SER 07.08.23 ...
                    TempDistrEntry."Location Code" := SalesInvoiceHeader."Location Code";
                    TempDistrEntry."Internal Company" := '';
                    // ... APO.002 SER 07.08.23
                end;
            Database::"Sales Cr.Memo Header":
                begin
                    RecRef.SetTable(SalesCrMemoHeader);
                    TempDistrEntry."Document Type" := Format(SalesHeader."Document Type"::"Credit Memo");
                    TempDistrEntry."Document No." := SalesCrMemoHeader."No.";
                    TempDistrEntry."Contact No." := SalesCrMemoHeader."Bill-to Contact No.";
                    TempDistrEntry."Customer No." := SalesCrMemoHeader."Bill-to Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := SalesCrMemoHeader."Bill-to Name";
                    TempDistrEntry."Language Code" := SalesCrMemoHeader."Language Code";
                    TempDistrEntry."Document Date" := SalesCrMemoHeader."Document Date";
                    // APO.002 SER 07.08.23 ...
                    TempDistrEntry."Location Code" := SalesCrMemoHeader."Location Code";
                    TempDistrEntry."Internal Company" := '';
                    // ... APO.002 SER 07.08.23
                end;
            Database::"Return Receipt Header":
                begin
                    RecRef.SetTable(ReturnReceiptHeader);
                    TempDistrEntry."Document Type" := Format(SalesHeader."Document Type"::"Return Order");
                    TempDistrEntry."Document No." := ReturnReceiptHeader."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := ReturnReceiptHeader."Bill-to Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := ReturnReceiptHeader."Bill-to Name";
                    TempDistrEntry."Language Code" := ReturnReceiptHeader."Language Code";
                    TempDistrEntry."Document Date" := ReturnReceiptHeader."Document Date";
                    // APO.002 SER 07.08.23 ...
                    TempDistrEntry."Location Code" := ReturnReceiptHeader."Location Code";
                    TempDistrEntry."Internal Company" := '';
                    // ... APO.002 SER 07.08.23
                end;
            Database::"Sales Header Archive":
                begin
                    RecRef.SetTable(SalesHeaderArch);
                    TempDistrEntry."Document Type" := Format(SalesHeaderArch."Document Type");
                    TempDistrEntry."Document No." := SalesHeaderArch."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := SalesHeaderArch."Bill-to Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := SalesHeaderArch."Bill-to Name";
                    TempDistrEntry."Language Code" := SalesHeaderArch."Language Code";
                    TempDistrEntry."Document Date" := SalesHeaderArch."Document Date";
                    // APO.002 SER 07.08.23 ...
                    TempDistrEntry."Location Code" := SalesHeaderArch."Location Code";
                    TempDistrEntry."Internal Company" := '';
                    // ... APO.002 SER 07.08.23
                end;
        end;
    end;

    local procedure GetFromFinancialDocument(RecRef: RecordRef; var TempDistrEntry: Record "Email Distribution Entry" temporary)
    var
        GenJournalLine: Record "Gen. Journal Line";
        BankAccountStatement: Record "Bank Account Statement";
        IssuedReminderHeader: Record "Issued Reminder Header";
    begin
        case RecRef.Number() of
            Database::"Issued Reminder Header":
                begin
                    RecRef.SetTable(IssuedReminderHeader);
                    TempDistrEntry."Document Type" := TextReminderType;
                    TempDistrEntry."Document No." := IssuedReminderHeader."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := IssuedReminderHeader."Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := IssuedReminderHeader.Name;
                    TempDistrEntry."Language Code" := IssuedReminderHeader."Language Code";
                    TempDistrEntry."Document Date" := IssuedReminderHeader."Document Date";
                end;
            Database::"Bank Account Statement":
                begin
                    RecRef.SetTable(BankAccountStatement);
                    TempDistrEntry."Document Type" := Format(BankAccountStatement."Statement Date");
                    TempDistrEntry."Document No." := BankAccountStatement."Statement No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := GetFirstSourceTableFieldCode(RecRef, Database::Customer);
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."Account No." := BankAccountStatement."Bank Account No.";
                    TempDistrEntry."To Name" := BankAccountStatement.TableCaption();
                    TempDistrEntry."Document Date" := BankAccountStatement."Statement Date";
                end;
            Database::"Gen. Journal Line":
                begin
                    RecRef.SetTable(GenJournalLine);
                    TempDistrEntry."Document Type" := Format(GenJournalLine."Account Type");
                    TempDistrEntry."Document No." := GenJournalLine."Document No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := GetFirstSourceTableFieldCode(RecRef, Database::Customer);
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."Account No." := GenJournalLine."Account No.";
                    TempDistrEntry."To Name" := TextCheckType;
                    TempDistrEntry."Document Date" := GenJournalLine."Document Date";
                end;
        end;
    end;

    local procedure GetFromServiceDocument(RecRef: RecordRef; var TempDistrEntry: Record "Email Distribution Entry" temporary)
    var
        ServiceHeader: Record "Service Header";
        ServContractHeader: Record "Service Contract Header";
        ServiceShptHeader: Record "Service Shipment Header";
        ServiceInvoiceHeader: Record "Service Invoice Header";
        ServiceCrMemoHeader: Record "Service Cr.Memo Header";
    begin
        case RecRef.Number() of
            Database::"Service Contract Header":
                begin
                    RecRef.SetTable(ServContractHeader);
                    TempDistrEntry."Document Type" := TextServiceContractType + Format(ServContractHeader."Contract Type");
                    TempDistrEntry."Document No." := ServContractHeader."Contract No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := ServContractHeader."Bill-to Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := ServContractHeader."Bill-to Name";
                    TempDistrEntry."Language Code" := ServContractHeader."Language Code";
                    TempDistrEntry."Document Date" := ServContractHeader."Starting Date";
                    TempDistrEntry."E-Mail" := ServContractHeader."E-Mail";
                    TempDistrEntry."Fax No." := ServContractHeader."Fax No.";
                end;
            Database::"Service Header":
                begin
                    RecRef.SetTable(ServiceHeader);
                    TempDistrEntry."Document Type" := TextServiceType + Format(ServiceHeader."Document Type");
                    TempDistrEntry."Document No." := ServiceHeader."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := ServiceHeader."Bill-to Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := ServiceHeader."Bill-to Name";
                    TempDistrEntry."Language Code" := ServiceHeader."Language Code";
                    TempDistrEntry."Document Date" := ServiceHeader."Document Date";
                    TempDistrEntry."E-Mail" := ServiceHeader."E-Mail";
                    TempDistrEntry."Fax No." := ServiceHeader."Fax No.";
                end;
            Database::"Service Shipment Header":
                begin
                    RecRef.SetTable(ServiceShptHeader);
                    TempDistrEntry."Document Type" := TextServiceShptType;
                    TempDistrEntry."Document No." := ServiceShptHeader."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := ServiceShptHeader."Bill-to Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := ServiceShptHeader."Bill-to Name";
                    TempDistrEntry."Language Code" := ServiceShptHeader."Language Code";
                    TempDistrEntry."Document Date" := ServiceShptHeader."Document Date";
                    TempDistrEntry."E-Mail" := ServiceShptHeader."E-Mail";
                    TempDistrEntry."Fax No." := ServiceShptHeader."Fax No.";
                end;
            Database::"Service Invoice Header":
                begin
                    RecRef.SetTable(ServiceInvoiceHeader);
                    TempDistrEntry."Document Type" := TextServiceInvoiceType;
                    TempDistrEntry."Document No." := ServiceInvoiceHeader."No.";
                    TempDistrEntry."Contact No." := ServiceInvoiceHeader."Bill-to Contact No.";
                    TempDistrEntry."Customer No." := ServiceInvoiceHeader."Bill-to Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := ServiceInvoiceHeader."Bill-to Name";
                    TempDistrEntry."Language Code" := ServiceInvoiceHeader."Language Code";
                    TempDistrEntry."Document Date" := ServiceInvoiceHeader."Document Date";
                    TempDistrEntry."E-Mail" := ServiceInvoiceHeader."E-Mail";
                    TempDistrEntry."Fax No." := ServiceInvoiceHeader."Fax No.";
                end;
            Database::"Service Cr.Memo Header":
                begin
                    RecRef.SetTable(ServiceCrMemoHeader);
                    TempDistrEntry."Document Type" := TextServiceCrMemoType;
                    TempDistrEntry."Document No." := ServiceCrMemoHeader."No.";
                    TempDistrEntry."Contact No." := ServiceCrMemoHeader."Bill-to Contact No.";
                    TempDistrEntry."Customer No." := ServiceCrMemoHeader."Bill-to Customer No.";
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."To Name" := ServiceCrMemoHeader."Bill-to Name";
                    TempDistrEntry."Language Code" := ServiceCrMemoHeader."Language Code";
                    TempDistrEntry."Document Date" := ServiceCrMemoHeader."Document Date";
                    TempDistrEntry."E-Mail" := ServiceCrMemoHeader."E-Mail";
                    TempDistrEntry."Fax No." := ServiceCrMemoHeader."Fax No.";
                end;
        end;
    end;

    local procedure GetFromOtherDocument(RecRef: RecordRef; var TempDistrEntry: Record "Email Distribution Entry" temporary)
    var
        AsmHeader: Record "Assembly Header";
        TransHeader: Record "Transfer Header";
        EmailDistrRecordColl_LT: Record "Email Distr. Record Coll.";
    begin
        case RecRef.Number() of
            Database::"Transfer Header":
                begin
                    RecRef.SetTable(TransHeader);
                    TempDistrEntry."Document Type" := TextTransferType;
                    TempDistrEntry."Document No." := TransHeader."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := GetFirstSourceTableFieldCode(RecRef, Database::Customer);
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."Transfer Code" := TransHeader."Transfer-from Code";
                    TempDistrEntry."To Name" := TransHeader."Transfer-from Name";
                    TempDistrEntry."Document Date" := TransHeader."Posting Date";
                end;
            Database::"Assembly Header":
                begin
                    RecRef.SetTable(AsmHeader);
                    TempDistrEntry."Document Type" := TextAssemblyType;
                    TempDistrEntry."Document No." := AsmHeader."No.";
                    TempDistrEntry."Contact No." := GetFirstSourceTableFieldCode(RecRef, Database::Contact);
                    TempDistrEntry."Customer No." := GetFirstSourceTableFieldCode(RecRef, Database::Customer);
                    TempDistrEntry."Vendor No." := GetFirstSourceTableFieldCode(RecRef, Database::Vendor);
                    TempDistrEntry."Item No." := AsmHeader."Item No.";
                    TempDistrEntry."To Name" := AsmHeader.Description;
                    TempDistrEntry."Document Date" := AsmHeader."Creation Date";
                end;
            Database::"Email Distr. Record Coll.":
                begin
                    RecRef.SetTable(EmailDistrRecordColl_LT);
                    TempDistrEntry."Document Type" := TextEmailDistrRecordColl;
                    TempDistrEntry."Document No." := Format(EmailDistrRecordColl_LT."No.");
                    TempDistrEntry."Contact No." := '';
                    TempDistrEntry."Customer No." := '';
                    TempDistrEntry."Vendor No." := '';
                    TempDistrEntry."Item No." := '';
                    TempDistrEntry."To Name" := '';
                    TempDistrEntry."Document Date" := DT2Date(EmailDistrRecordColl_LT."Creation Date/Time");
                end;
        end;
    end;

    local procedure "--- Help Ftcs. ---"()
    begin
    end;

    procedure GetFirstSourceTableFieldCode(RecordVariant: Variant; RelationTableNo: Integer): Text
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

    local procedure SetResponsibilityCenter(RecRef: RecordRef; var TempDistrEntry: Record "Email Distribution Entry" temporary): Text
    begin
        if TempDistrEntry."Resp. Center Code" = '' then
            TempDistrEntry."Resp. Center Code" := GetFirstSourceTableFieldCode(RecRef, Database::"Responsibility Center");
    end;

    local procedure SetEmailAndFax(var TempDistrEntry: Record "Email Distribution Entry" temporary): Text
    var
        Customer: Record Customer;
        Vendor: Record Vendor;
        ShiptoAddress_LT: Record "Ship-to Address";
        OrderAddress_LT: Record "Order Address";
        Contact: Record Contact;
    begin
        if ((TempDistrEntry."Vendor No." <> '') and (TempDistrEntry."Order Address Code" = '')) or
            ((TempDistrEntry."Customer No." <> '') and (TempDistrEntry."Ship-to Address Code" = '')) then begin
            if Contact.Get(TempDistrEntry."Contact No.") then begin
                if TempDistrEntry."E-Mail" = '' then
                    TempDistrEntry."E-Mail" := Contact."E-Mail";
                if TempDistrEntry."Fax No." = '' then
                    TempDistrEntry."Fax No." := Contact."Fax No.";
                exit;
            end;
        end;

        if Customer.Get(TempDistrEntry."Customer No.") then begin
            if TempDistrEntry."Ship-to Address Code" <> '' then begin
                if ShiptoAddress_LT.Get(Customer."No.", TempDistrEntry."Ship-to Address Code") then begin
                    TempDistrEntry."E-Mail" := ShiptoAddress_LT."E-Mail";
                    TempDistrEntry."Fax No." := ShiptoAddress_LT."Fax No.";
                end;
            end;
            if TempDistrEntry."E-Mail" = '' then
                TempDistrEntry."E-Mail" := Customer."E-Mail";
            if TempDistrEntry."Fax No." = '' then
                TempDistrEntry."Fax No." := Customer."Fax No.";
            exit;
        end;

        if Vendor.Get(TempDistrEntry."Vendor No.") then begin
            if TempDistrEntry."Order Address Code" <> '' then begin
                if OrderAddress_LT.Get(Vendor."No.", TempDistrEntry."Order Address Code") then begin
                    TempDistrEntry."E-Mail" := OrderAddress_LT."E-Mail";
                    TempDistrEntry."Fax No." := OrderAddress_LT."Fax No.";
                end;
            end;
            if TempDistrEntry."E-Mail" = '' then
                TempDistrEntry."E-Mail" := Vendor."E-Mail";
            if TempDistrEntry."Fax No." = '' then
                TempDistrEntry."Fax No." := Vendor."Fax No.";
            exit;
        end;
    end;
}

