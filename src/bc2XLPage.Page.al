/// <summary>
/// Page PageWithAddIn (ID 50130).
/// </summary>
page 72000 PageWithAddIn
{
    UsageCategory=Administration;
    ApplicationArea=all;
    caption = 'BC2Excel';
    
    layout
    {
        area(Content)
        {
             // The control add-in can be placed on the page using usercontrol keyword.
            usercontrol(AddInn; BC2XL)
            {

                ApplicationArea = All;

                 // The control add-in events can be handled by defining a trigger with a corresponding name.
                trigger GetWorksheetAsJson(d: JsonObject)
                begin
                   // message('from JS: %1', s); 
                   message('data', d);
                end;
            }
        }
    }

    actions
    {
        area(Reporting)
        {
            action(CreateDoc)
            {
                ApplicationArea = All;
                caption = 'Excel Dokument generieren';
                
                trigger OnAction();
                var 
                   docatt : record "Document Attachment";
                   Customer: Record customer;
                //   media : record "Tenant Media";
                //   jsonData: jsonObject;
                   Ins: Instream;
                   Outs: Outstream;
                //   Content: text;
                Base64: codeunit "Base64 Convert";
                Base64Content: Text;
                //   uri: codeunit DotNet_BinaryWriter;
                //   mstr: codeunit DotNet_MemoryStream;
                   tempBlob: codeunit "temp Blob";
                data: JsonObject;
                
                //jstk: jsonToken;
                begin
                    docatt.SETRANGE("Table ID",database::customer);
                    docatt.SETRANGE("No.", '01445544');
                    customer.GET('01445544');
                    if docatt.FINDSET(FALSE, FALSE) then BEGIN                    
                       begin
                        tempBlob.CreateOutStream(outs);  
                        docatt."Document Reference ID".ExportStream(outs);
                        tempblob.CreateInStream(ins);
                        Base64Content:=base64.toBase64(ins);

                        // data:=StrSubstNo('{first_name: "%1",last_name: "%2",phone: "%3",description: "%4"}',customer.Name,customer."Name 2",customer."Phone No.",customer."Home Page");
                        
                        data.Add('Firmenname', Customer.Name);
                        data.Add('Adresse', Customer.Address);
                        data.Add('PLZ', Customer."Post Code");
                        data.Add('Ort', Customer.City);
                        data.Add('Datum', Normaldate(Workdate));
                        //  IF media.get(docatt."Document Reference ID".MediaID()) then BEGIN                         
                        //     Media.CALCFIELDS(Content);
                        //     Media.Content.CreateInStream(ins);
                        //     ins.Readtext(content);
                        //     Base64Content:=base64.toBase64(content);
                        //  end;  
                       end;

                    end;

                    // The control add-in methods can be invoked via a reference to the usercontrol.
                    CurrPage.AddInn.readXLSX(Base64Content);
                end;

            }

             action(CreateExcelFromMediaField)
                {
                    ApplicationArea = All;
                    Caption = 'Create Excel from given Media Field';
                    
                    trigger OnAction()
                    var
                      bc2xlMgmt: codeunit "BC2XL Management";
                      docatt: record "Document Attachment";
                      refdocatt: RecordRef;

                    begin
                        docatt.SETRANGE("Table ID",database::customer);
                        docatt.SETRANGE("No.", '01445544');
                        if docatt.findset then BEGIN
                          refdocatt.GetTable(docatt);
                          bc2xlMgmt.setAddIn(currpage.AddInn);
                          bc2xlMgmt.getFileFromMediaField(refdocatt,docatt.fieldno("Document Reference ID"));
                          bc2xlMgmt.createXLSX('testout.xlxs');
                        END; 
                    end;
                }
        }
    }


   Procedure GetToken(Key_L: Text; JsonObject_L: JsonObject): Text
   var
        JsonToken_L: JsonToken;
        Value_L: Text;
    begin
        Clear(JsonToken_L);
        Clear(Value_L);
        JsonObject_L.GET(Key_L,JsonToken_L);
        JsonToken_L.WriteTo(Value_L);
        IF ((StrPOS(Value_L,'"') <>0 )) THEN begin
            Value_L:=CopyStr(Value_L,2,strLEN(Value_L)-2);
            exit(Value_L);
        end;
        IF Value_L='null' then 
            Value_L:='';
        IF Value_L='' then 
            Value_L:='';
        exit(Value_L);    
    end;    

}