codeunit 50990 "Nachpflege APO"
{
    // version ApoDE14.71
    Permissions = tabledata "Item Ledger Entry" = rm,
                  tabledata "Sales Shipment Header" = rm,
                  tabledata "Sales Shipment Line" = rm,
                  tabledata "Sales Invoice Header" = rm,
                  tabledata "Sales Invoice Line" = rm,
                  tabledata "Sales Cr.Memo Header" = rm,
                  tabledata "Sales Cr.Memo Line" = rm,
                  tabledata "Purch. Rcpt. Header" = rm,
                  tabledata "Purch. Inv. Header" = rm,
                  tabledata "Return Receipt Header" = rm,
                  tabledata "Return Receipt Line" = rm;

    trigger OnRun()
    begin
        //!!!!!!!!!!!!!!!!!!!!!!!!!!!
        if not (UserId() in ['APODISCOUNTER\M.THOMANN', 'APODISCOUNTER\DEV.M.THOMANN', 'APODISCOUNTER\S.ERDMENGER', 'APODISCOUNTER\DEV.S.ERDMENGER', 'APODISCOUNTER\NAV_SVCACC', 'APODISCOUNTER\J.UNGER', 'APODISCOUNTER\DEV.J.UNGER']) then
            exit;

        CheckDatabase();
        //!!!!!!!!!!!!!!!!!!!!!!!!!!!

        UpdateB2BCustomer();

        Message(ProcessingFinishedMsgTxt_G);
    end;

    var
        ExcelBuffer_GT: Record "Excel Buffer" temporary;
        ApoMasterDataSetup_GT: Record "Apo Master Data Setup";
        DialogMgt_GC: Codeunit "Dialog Mgt.";
        DoYouWantStartQstTxt_G: Label 'Do you want start processing?';
        ExcelFilterString: Label 'Excel Files (*.xlsx)|*.xlsx';
        ImportFileCaption: Label 'Import File';
        ProcessingFinishedMsgTxt_G: Label 'Processing finished';

    local procedure _FunctionTemplate()
    var
        DummyRec: Record "Integer";
        DataTypeManagement: Codeunit "Data Type Management";
        RecRef: RecordRef;
        FldRef: FieldRef;
        Window_L: Dialog;
        i: Integer;
        j: Integer;
        k: Integer;
        TotalRows_L: Integer;
        ProgressBarDlgTxt: Label 'Progress #1###### #2######';
    begin

        if not Confirm(DoYouWantStartQstTxt_G) then
            exit;

        DummyRec.Reset();
        DummyRec.SetRange(Number, 1, 10);
        TotalRows_L := DummyRec.Count();
        Window_L.Open(ProgressBarDlgTxt);
        j := 0;
        i := 0;
        if DummyRec.FindSet(true, false) then begin
            repeat
                Window_L.Update(1, DummyRec.Number);
                i += 1;
                j += 1;
                if j = 1000 then begin
                    Commit();
                    j := 0;
                end;
                Window_L.Update(2, StrSubstNo('%1 / %2', Format(i), Format(TotalRows_L)));
            until DummyRec.Next() = 0;
        end;
        Window_L.Close();
    end;

    local procedure CheckDatabase()
    var
        Session_LT: Record Session;
    begin

        Session_LT.Reset();
        Session_LT.SetRange("User ID", UserId());
        Session_LT.SetRange("My Session", true);
        Session_LT.FindFirst();

        if not Confirm(CompanyName() + '\ \' + UpperCase(Session_LT."Database Name") + '\ \', false) then
            Error('');
    end;

    local procedure UpdateB2BCustomer()
    var
        Customer_LT: Record Customer;
        DataTypeManagement: Codeunit "Data Type Management";
        RecRef: RecordRef;
        FldRef: FieldRef;
        Window_L: Dialog;
        i: Integer;
        j: Integer;
        k: Integer;
        TotalRows_L: Integer;
        ProgressBarDlgTxt: Label 'Progress #1###### #2######';
    begin

        if not Confirm(DoYouWantStartQstTxt_G) then
            exit;

        Customer_LT.Reset();
        Customer_LT.SetRange("Responsibility Center", 'SAL_NLB2B');
        Customer_LT.SetRange("Document Sending Profile", '');
        Customer_LT.ModifyAll("Document Sending Profile", 'E-MAIL');
    end;
}

