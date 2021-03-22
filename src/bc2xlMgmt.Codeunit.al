

/// <summary>
/// Codeunit BC2XL Management (ID 50100).
/// </summary>
codeunit 72000 "BC2XL Management"
{


    var
        AddIn: controladdin BC2XL;
        base64Content: text;
        jsonData: JsonObject;

    /// <summary>
    /// setAddIn.
    /// </summary>
    /// <param name="_AddIn">VAR ControlAddIn.</param>
    procedure setAddIn(_AddIn: ControlAddIn BC2XL)
    begin
        AddIn := _Addin;
    end;

    
    Procedure getFileFromMediaField(recRef: recordRef; fieldID: integer) fileFound: boolean
    var
        mediaFieldRef: FieldRef;
        mediaID: guid;
        tempBlob: codeunit "Temp blob";
        tenantMedia: Record "Tenant Media";
        tenantMediaref: RecordRef;
        ins: instream;
        base64: codeunit "Base64 Convert";
        blobFieldRef: FieldRef;
    begin
        //Record Ref auf das Media Feld dort erhalte ich die MEDIAID 
        fileFound:=FALSE;
        mediaFieldRef := recRef.field(FieldID);
        mediaID := mediaFieldref.Value;
        
        //Im Feld ID finde ich die gesuchte MEDIAID. Von dort hole ich mir den Content aus dem Feld Content
        //Über die Codeunit TempBlob erhalte ich den benötigten InStream um auf BAse64 zu konvertieren
        tenantMedia.SETRANGE(ID, mediaID);
        IF tenantMedia.FINDSET() then BEGIN
            tenantMedia.CALCFIELDS(Content);
            tenantMediaref.GetTable(tenantMedia);
            blobFieldRef := tenantMediaref.FIELD(tenantMedia.Fieldno(Content));
            tempblob.FromFieldRef(blobFieldRef);
            tempblob.CreateInStream(ins);
            Base64Content := base64.toBase64(ins);
            FileFound:=TRUE;
        END;

    end;

    /// <summary>
    /// readXLSX.
    /// </summary>
    procedure readXLSX()
    var
      c: codeunit "OAuth 2.0 Mgt.";
      r: Record "OAuth 2.0 Setup";
    begin
        // c.RefreshAccessToken();
        // r.
       Addin.readXLSX(base64Content); 
    end;

    /// <summary>
    /// create XLSX-File
    /// </summary>
    /// <param name="Filename">text.</param>
    procedure createXLSX(Filename: text)
    begin
        Addin.createXLSX(Filename);
    end;


}