

// codeunit 72001 ExcelVerwalt
// {

//     trigger OnRun()
//     begin
//     end;

//     var
       
//         // wb: Dotnet xlWorkbook;
//         // ws: dotnet xlWorksheet;
//         // rg: DotNet xlRange;
//         // rg2: DotNet xlRange;
//         // col: Integer;
//         // Shape: dotnet xlShape;
//         // Hoehe: Decimal;
//         // pb: dotnet xlHPagebreaks;
//         // rows: dotnet xlrange;
//         // wss: dotnet xlWorksheets;
//         // ProjEinr: Record "Jobs Setup";
//         // ZeileProSeite: Integer;
//         // wd: dotnet xlWindow;
//         // xlBorder: dotnet xlBorders;
//         // ActChart: dotnet xl_Chart;
//         // DataRange: dotnet xlrange;
//         // window: dotnet xlWindow;
//         // xlLine: dotnet xlShape;
//         // xlLineFormat: dotnet xlLineformat;
//         // xlColor: dotnet xlColorFormat;
//         // xlShapes: dotnet xlShapes;
//         // i: Integer;
//         // sel: dotnet xlShapes;
//         // txtVerteiler: Text[250];
//         // Filemgmt: Codeunit "File Management";
//         // FileMgmt2: Codeunit "FileManagement/ow";
//         // AktChart: dotnet xlChart;
//         // colCharts: dotnet xlsheets;

//     //[Scope('OnPrem')]
//     procedure AppInit(visible: Boolean)
//     begin
//         // AppInit()

//         // Vorherige Instanz löschen
//         Clear(xlApp);

//         // Neues Excel Instanzieren
//         //if not Create(xlApp, true, true) then
//         xlApp := xlapp.Application;
//         xlapp.Visible(visible);
//     end;
//     //[Scope('OnPrem')]
//     procedure AppInit2()
//     begin
//         // AppInit2()

//         appInit(true);
//         ws := xlApp.ActiveWorkbook.ActiveSheet;
//     end;

//     //[Scope('OnPrem')]
//     procedure AppClose()
//     begin
//         // AppClose()
//         Clear(xlApp);
//     end;

//     //[Scope('OnPrem')]
//     procedure AppCloseSilent()
//     begin
//         xlApp.DisplayAlerts(false);
//         xlApp.Quit;
//         Clear(xlApp);
//     end;

//     //[Scope('OnPrem')]
//     procedure Setvisible(bvisible: Boolean)
//     begin
//         xlApp.Visible(bvisible);
//     end;

//     //[Scope('OnPrem')]
//     procedure NewWorkbook()
//     begin
//         // NewWorkbook

//         wb := xlApp.Workbooks.Add;
//     end;

//     //[Scope('OnPrem')]
//     procedure GetActiveWorkbook()
//     begin
//         wb := xlApp.ActiveWorkbook;
//     end;

//     //[Scope('OnPrem')]
//     procedure AddXLT(vFilename: Text[250]): Boolean
//     begin
//         if Filemgmt.ClientFileExists(vFilename) then
//             wb := xlApp.Workbooks.Add(vFilename);
//         exit(Filemgmt.ClientFileExists(vFilename));
//     end;

//     //[Scope('OnPrem')]
//     procedure OpenWorkbook(Dateiname: Text[250])
//     begin
//         // OpenWorkbook

//         if Dateiname = '' then begin

//             // Dialogbox öffnen

//         end else begin

//             // Workbook öffnen
//             wb := xlApp.Workbooks.Open(Dateiname);


//         end;
//     end;

//     //[Scope('OnPrem')]
//     procedure SaveWorkbook()
//     begin
//         // SaveWorkbook

//         wb.Save;
//     end;

//     //[Scope('OnPrem')]
//     procedure SaveWorkbookAs(Path: Text[250])
//     begin
//         wb.SaveAs(Path);
//         //wb.SaveCopyAs(Path);
//     end;

//     //[Scope('OnPrem')]
//     procedure NewWorksheet()
//     begin
//         // NewWorksheet

//         //ws :=wb.Worksheets.Add(,After,1,-4167);
//         ws := wb.Worksheets.Add
//     end;

//     //[Scope('OnPrem')]
//     procedure NewWorksheetByName(Sheetname: Text[250])
//     begin
//         // NewWorksheetByName
//         ws := wb.Worksheets.Add;
//         ws.Name(Sheetname);
//     end;

//     //[Scope('OnPrem')]
//     procedure SetWorksheet(index: Integer; Sheetname: Text[250])
//     begin
//         // SetWorksheet

//         ws := wb.Worksheets.Item(index);

//         if Sheetname <> '' then
//             ws.Name(Sheetname);

//         ws.Activate;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetWorksheetByName(Sheetname: Text[250])
//     var
//         wstemp: dotnet xlWorksheet;
//         i: Integer;
//     begin
//         for i := 1 to wb.Worksheets.Count do begin
//             wstemp := wb.Worksheets.Item(i);
//             if wstemp.Name = Sheetname then begin
//                 ws := wb.Worksheets.Item(i);
//                 ws.Activate;
//             end;
//         end;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetValueOnCell(col: Integer; row: Integer; Value: Text[250])
//     var
//         xlCells: dotnet xlRange;
//     begin
//         // SetValueOnCell

//         ws.Range(ColLetter(col) + Format(row)).Value(Value);
//     end;

//     //[Scope('OnPrem')]
//     procedure GetValueOnCell(col: Integer; row: Integer) value: Text[250]
//     begin
//         // GetValueOnCell

//         exit(Format(ws.Range(ColLetter(col) + Format(row)).Value));
//     end;

//     //[Scope('OnPrem')]
//     procedure SetFormatOnCell(col: Integer; row: Integer; Formatstring: Text[250])
//     begin
//         ws.Range(ColLetter(col) + Format(row)).NumberFormat(Formatstring);
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellBold(col: Integer; row: Integer)
//     begin
//         // SetCellBold

//         ws.Range(ColLetter(col) + Format(row)).Font.Bold(true);
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellAlignRight(col: Integer; row: Integer)
//     begin
//         ws.Range(ColLetter(col) + Format(row)).HorizontalAlignment := -4152; //xlRight
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellHorizAlign(col: Integer; row: Integer; position: Integer)
//     var
//         pos: Integer;
//     begin
//         case position of
//             0:
//                 pos := -4131;  // xlleft
//             1:
//                 pos := -4108;  // xlcenter
//             2:
//                 pos := -4152;  // xlright
//         end;
//         ws.Range(ColLetter(col) + Format(row)).HorizontalAlignment := pos;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellOrientation(col: Integer; row: Integer; angle: Decimal)
//     begin
//         ws.Range(ColLetter(col) + Format(row)).Orientation := angle;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellWrapText(col: Integer; row: Integer; boolWrapText: Boolean)
//     begin
//         ws.Range(ColLetter(col) + Format(row)).WrapText := boolWrapText;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellFontSize(col: Integer; row: Integer; vsize: Integer)
//     begin
//         ws.Range(ColLetter(col) + Format(row)).Font.Size := vsize;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellVertAlign(col: Integer; row: Integer; optPosition: Option "top;center;bottom")
//     var
//         iposition: Integer;
//     begin
//         // SetCellVertAlign

//         case optPosition of
//             0:
//                 iposition := -4160;
//             1:
//                 iposition := -4108;
//             2:
//                 iposition := -4107;
//         end;

//         ws.Range(ColLetter(col) + Format(row)).VerticalAlignment := iposition; //xltop
//     end;

//     //[Scope('OnPrem')]
//     procedure SetBorder(col: Integer; row: Integer; xcol: Integer; xrow: Integer; Thick: Boolean)
//     begin
//         // Setborder
//         case Thick of

//             false:
//                 begin
//                     if (xcol = 0) and (xrow = 0) then
//                         ws.Range(ColLetter(col) + Format(row)).BorderAround(1, 2, -4105, 0)
//                     else
//                         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).BorderAround(1, 1, -4105, 0)
//                 end;

//             true:
//                 begin
//                     if (xcol = 0) and (xrow = 0) then
//                         ws.Range(ColLetter(col) + Format(row)).BorderAround(1, 4, -4105, 0)
//                     else
//                         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).BorderAround(1, 2, -4105, 0)
//                 end;

//         end;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetMultiBorder(col: Integer; row: Integer; xcol: Integer; xrow: Integer)
//     var
//         i: Integer;
//     begin
//         // SetMultiBorder

//         for i := 7 to 12 do begin
//             ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).Borders.Item(i).LineStyle := 1;
//             ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).Borders.Item(i).Weight := 1; //xlHairline
//         end;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetBorderLine(col: Integer; row: Integer; xcol: Integer; xrow: Integer; which: Integer; intWeight: Integer)
//     begin
//         // SetBorderLine
//         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).Borders.Item(which).LineStyle := 1;
//         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).Borders.Item(which).Weight := intWeight; //xlHairline
//     end;

//     //[Scope('OnPrem')]
//     procedure SetmultiBorderLine(col: Integer; row: Integer; xcol: Integer; xrow: Integer)
//     var
//         i: Integer;
//     begin
//         // SetMultiBorderLine

//         for i := 7 to 11 do begin
//             ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).Borders.Item(i).LineStyle := 1;
//         end;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetRowWidth(col: Integer; Faktor: Decimal)
//     begin
//         // SetRowWidth()

//         ws.Range(ColLetter(col) + ':' + ColLetter(col)).ColumnWidth(Faktor);
//     end;

//     //[Scope('OnPrem')]
//     procedure SetRowHeight(row: Integer; Faktor: Decimal)
//     begin
//         // SetRowHeight()

//         ws.Range(Format(row) + ':' + Format(row)).RowHeight(Faktor);
//     end;

//     //[Scope('OnPrem')]
//     procedure FormatCell(col: Integer; row: Integer; Value: Text[250])
//     begin
//         // FormatCell()

//         //ws.Range(ColLetter(col)+FORMAT(row))
//     end;

//     //[Scope('OnPrem')]
//     procedure ColLetter(col: Integer): Text[30]
//     var
//         Buchstabe: Text[30];
//         PosBuchstabe1: Integer;
//         PosBuchstabe2: Integer;
//         Offset: Integer;
//     begin
//         // ColLetter

//         Offset := 64;
//         PosBuchstabe1 := Round(col / 26, 1, '>') - 1;
//         PosBuchstabe2 := col mod 26;
//         if PosBuchstabe2 = 0 then
//             PosBuchstabe2 := 26;

//         if PosBuchstabe1 = 0 then begin
//             Buchstabe[1] := Offset + PosBuchstabe2;
//             Buchstabe[2] := ' ';
//         end else begin
//             Buchstabe[1] := Offset + PosBuchstabe1;
//             Buchstabe[2] := Offset + PosBuchstabe2;
//         end;

//         Buchstabe := DelChr(Buchstabe, '=', ' ');

//         exit(Buchstabe);
//     end;

//     //[Scope('OnPrem')]
//     procedure GetPictureFromStream(var instr: InStream; x: Decimal; y: Integer; width: Decimal; height: Decimal)
//     var
//         range: dotnet xlRange;
//         h: Decimal;
//     begin
//         // GetPictureFromStream()

//         h := ws.Range('A1:A' + Format(y)).Height;
//         xlShapes := ws.Shapes;
//         //Shape := ws.Shapes.AddPicture(Pfad, 0, -1, x,h, width, height);
//         //shape:=xlShapes.AddShape(1, x,h, width, height);
//     end;

//     //[Scope('OnPrem')]
//     procedure GetPicture(x: Decimal; y: Integer; Pfad: Text[250]; width: Decimal; height: Decimal)
//     var
//         range: dotnet xlRange;
//         h: Decimal;
//     begin
//         // GetPicture()

//         h := ws.Range('A1:A' + Format(y)).Height;
//         xlShapes := ws.Shapes;
//         //Shape := ws.Shapes.AddPicture(Pfad, 0, -1, x,h, width, height);
//         xlShapes.AddPicture(Pfad, 0, -1, x, h, width, height);
//     end;

//     //[Scope('OnPrem')]
//     procedure SetPrinter()
//     begin
//         // SetPrinter()
//         ProjEinr.Get;
//         ws.PageSetup.Orientation := 2; //Landscape
//         //ws.PageSetup.FitToPagesWide := 1;
//         //ws.PageSetup.FitToPagesTall := 200;
//         //ws.PageSetup.CenterHeader:='&6Verteiler:'+ProjEinr."Verteiler Wochenliste";
//         ws.PageSetup.CenterFooter := 'Seite &S';
//         ws.PageSetup.RightFooter := '&D';

//         ws.PageSetup.LeftMargin := xlApp.InchesToPoints(0.27559);
//         ws.PageSetup.RightMargin := xlApp.InchesToPoints(0.27559);
//         ws.PageSetup.TopMargin := xlApp.InchesToPoints(0.826771653);
//         ws.PageSetup.BottomMargin := xlApp.InchesToPoints(0.47244);
//         ws.PageSetup.HeaderMargin := xlApp.InchesToPoints(0.51181);
//         ws.PageSetup.FooterMargin := xlApp.InchesToPoints(0.23622);

//         ws.PageSetup.FirstPageNumber := -4105; //xlAutomatic
//         ws.PageSetup.Zoom := 100;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetPrinterTermine(Projekt: Record Job)
//     begin
//         // SetPrinterTermine()

//         //ws.PageSetup.Orientation:=2; //Landscape
//         //ws.PageSetup.FitToPagesWide := 1;
//         //ws.PageSetup.FitToPagesTall := 200;
//         //ws.PageSetup.CenterHeader:='&6Verteiler: '+txtVerteiler;
//         ws.PageSetup.CenterFooter := '&6' + Projekt.Description;
//         ws.PageSetup.LeftFooter := '&6Ausdruck vom &D';
//         ws.PageSetup.RightFooter := '&6Seite &S';

//         ws.PageSetup.LeftMargin := xlApp.InchesToPoints(0.27559);
//         ws.PageSetup.RightMargin := xlApp.InchesToPoints(0.27559);
//         ws.PageSetup.TopMargin := xlApp.InchesToPoints(0.826771653);
//         ws.PageSetup.BottomMargin := xlApp.InchesToPoints(0.47244);
//         ws.PageSetup.HeaderMargin := xlApp.InchesToPoints(0.51181);
//         ws.PageSetup.FooterMargin := xlApp.InchesToPoints(0.23622);

//         ws.PageSetup.FirstPageNumber := -4105; //xlAutomatic
//         ws.PageSetup.Zoom := 100;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetPrinterFertigungsstand()
//     begin
//         // SetPrinterFertigungsstand()

//         //ws.PageSetup.Orientation:=2; //Landscape
//         //ws.PageSetup.FitToPagesWide := 1;
//         //ws.PageSetup.FitToPagesTall := 200;
//         //ws.PageSetup.CenterHeader:='&6Verteiler:'+ProjEinr."Verteiler Wochenliste";
//         //ws.PageSetup.CenterFooter:='&6Verteiler: TL, PM, EK, AV, Ko, 2D/3D, Ltg. WB, Mst. WB 2x, Akte, Projekt';
//         ws.PageSetup.LeftFooter := '&6Verteiler: ' + txtVerteiler;
//         //ws.PageSetup.RightFooter:='&6Seite &S';

//         ws.PageSetup.LeftMargin := xlApp.InchesToPoints(0.19685);
//         ws.PageSetup.RightMargin := xlApp.InchesToPoints(0.19685);
//         ws.PageSetup.TopMargin := xlApp.InchesToPoints(0.7874);
//         ws.PageSetup.BottomMargin := xlApp.InchesToPoints(0.47244);
//         ws.PageSetup.HeaderMargin := xlApp.InchesToPoints(0.51181);
//         ws.PageSetup.FooterMargin := xlApp.InchesToPoints(0.23622);

//         ws.PageSetup.FirstPageNumber := -4105; //xlAutomatic
//         ws.PageSetup.Zoom := 115;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetPrinterQMPlan()
//     begin
//         // SetPrinter()
//         ProjEinr.Get;
//         //ws.PageSetup.FitToPagesWide := 1;
//         //ws.PageSetup.FitToPagesTall := 200;
//         ws.PageSetup.CenterFooter := 'Seite &S';
//         ws.PageSetup.RightFooter := '&D';

//         ws.PageSetup.LeftMargin := xlApp.InchesToPoints(0.27559);
//         ws.PageSetup.RightMargin := xlApp.InchesToPoints(0.27559);
//         ws.PageSetup.TopMargin := xlApp.InchesToPoints(0.826771653);
//         ws.PageSetup.BottomMargin := xlApp.InchesToPoints(0.47244);
//         ws.PageSetup.HeaderMargin := xlApp.InchesToPoints(0.51181);
//         ws.PageSetup.FooterMargin := xlApp.InchesToPoints(0.23622);

//         ws.PageSetup.FirstPageNumber := -4105; //xlAutomatic
//         ws.PageSetup.Zoom := 100;
//         ws.PageSetup.PrintTitleRows := '$1:$6';
//     end;

//     //[Scope('OnPrem')]
//     procedure SetPrinterLieferabruf()
//     begin
//         // SetPrinter()
//         ProjEinr.Get;
//         ws.PageSetup.Orientation := 2; //Landscape
//         //ws.PageSetup.FitToPagesWide := 1;
//         //ws.PageSetup.FitToPagesTall := 200;

//         //ws.PageSetup.CenterHeader:='&10Verteiler:'+ProjEinr."Verteiler Lieferabrufe";

//         ws.PageSetup.CenterFooter := 'Seite &S';
//         ws.PageSetup.RightFooter := 'Stand: &D';

//         ws.PageSetup.LeftMargin := xlApp.InchesToPoints(0.27559);
//         ws.PageSetup.RightMargin := xlApp.InchesToPoints(0.27559);
//         ws.PageSetup.TopMargin := xlApp.InchesToPoints(0.826771653);
//         ws.PageSetup.BottomMargin := xlApp.InchesToPoints(0.47244);
//         ws.PageSetup.HeaderMargin := xlApp.InchesToPoints(0.51181);
//         ws.PageSetup.FooterMargin := xlApp.InchesToPoints(0.23622);

//         ws.PageSetup.FirstPageNumber := -4105; //xlAutomatic
//         ws.PageSetup.Zoom := 100;
//         ws.PageSetup.PrintTitleRows := '$1:$6';
//     end;

//     //[Scope('OnPrem')]
//     procedure SetHPagesBreaks()
//     var
//         "max": Integer;
//         i: Integer;
//         LetzterProjektAnfang: Integer;
//         Umbruch: Integer;
//         MaxUmbruch: Integer;
//         rgText: Text[200];
//         ReiheUmbruch: Integer;
//         ReiheUmbruchtxt: Text[200];
//     begin
//         exit;
//         // SetHPagesBreaks
//         LetzterProjektAnfang := 1;
//         xlApp.ActiveWindow.View := 2; //xlPagebreakView
//         MaxUmbruch := ws.HPageBreaks.Count;
//         if MaxUmbruch = 0 then exit;

//         Umbruch := 1;
//         ZeileProSeite := 1;


//         rg := ws.UsedRange;
//         max := rg.Rows.Count;
//         for i := 2 to max do begin
//             ZeileProSeite += 1;
//             rg := ws.Range(ColLetter(1) + Format(i));
//             rgText := Format(rg.Value);
//             if rgText = 'Projekt' then
//                 LetzterProjektAnfang := i;
//             if (ZeileProSeite > 38) and
//                (rgText <> 'Projekt') then begin
//                 //ws.HPageBreaks.Item(Umbruch).Delete;
//                 ws.HPageBreaks.Add(ws.Range(ColLetter(1) + Format(LetzterProjektAnfang)));
//                 Umbruch := Umbruch + 1;
//                 ZeileProSeite := 0;
//             end;
//         end;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellPattern(col: Integer; row: Integer; color: Integer)
//     begin
//         // SetCellpattern   -4108 - xlCenter

//         ws.Range(ColLetter(col) + Format(row)).Interior.ColorIndex := color;
//         ws.Range(ColLetter(col) + Format(row)).Interior.Pattern := 1; //xlsolid
//         ws.Range(ColLetter(col) + Format(row)).HorizontalAlignment := -4108;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellMultiPattern(col: Integer; row: Integer; xcol: Integer; xrow: Integer; color: Integer)
//     begin
//         // SetCellMultiPattern
//         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).Interior.ColorIndex := color;
//         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).Interior.Pattern := 1; //xlsolid
//         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).HorizontalAlignment := -4108;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellFontColor(col: Integer; row: Integer; color: Integer)
//     begin
//         //color 3 = rot
//         ws.Range(ColLetter(col) + Format(row)).Font.ColorIndex := color;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellMultiFontColor(col: Integer; row: Integer; xcol: Integer; xrow: Integer; color: Integer)
//     begin
//         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).Font.ColorIndex := color;
//     end;

//     //[Scope('OnPrem')]
//     procedure Freeze(col: Integer; row: Integer)
//     begin
//         // Freeze

//         ws.Range(ColLetter(col) + Format(row)).Select;
//         wd := xlApp.ActiveWindow;
//         wd.FreezePanes(true);
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellToCenter(col: Integer; row: Integer)
//     begin
//         // SetCellToCenter

//         ws.Range(ColLetter(col) + Format(row)).HorizontalAlignment := -4108;
//     end;

//     //[Scope('OnPrem')]
//     procedure MergeCells(col: Integer; row: Integer; xcol: Integer; xrow: Integer; Center: Boolean)
//     begin
//         // MergeCells

//         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).Merge;

//         if Center then
//             ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).HorizontalAlignment := -4108; //xlcenter
//     end;

//     //[Scope('OnPrem')]
//     procedure MergeCellsRight(col: Integer; row: Integer; xcol: Integer; xrow: Integer)
//     begin
//         // MergeCellsRight

//         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).Merge;

//         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).HorizontalAlignment := -4152; //xlRight
//     end;

//     //[Scope('OnPrem')]
//     procedure GetCellColor(col: Integer; row: Integer): Integer
//     begin
//         exit(ws.Range(ColLetter(col) + Format(row)).Interior.ColorIndex);
//     end;

//     //[Scope('OnPrem')]
//     procedure AddChart(Charttype: Integer; PlotBy: Integer; x1: Integer; y1: Integer; x2: Integer; y2: Integer; ShowTable: Boolean; Chartname: Text[80])
//     begin
//         // AddChart

//         ActChart := xlApp.Charts.Add;
//         ActChart.ChartType := Charttype;
//         DataRange := ws.Range(ColLetter(x1) + Format(y1) + ':' + ColLetter(x2) + Format(y2));
//         ActChart.SetSourceData(DataRange, PlotBy);
//         ActChart.Location(1, Chartname);
//         ActChart.HasDataTable := ShowTable;
//         ActChart.DataTable.ShowLegendKey := true;
//     end;

//     //[Scope('OnPrem')]
//     procedure NullWerteAnzeigen("unterdrücken": Boolean)
//     begin
//         window := xlApp.ActiveWindow;
//         window.DisplayZeros(unterdrücken);
//     end;

//     //[Scope('OnPrem')]
//     procedure HorizontalesBlatt()
//     begin
//         ws.PageSetup.Orientation := 2
//     end;

//     //[Scope('OnPrem')]
//     procedure SetPrinterPortrait(Deb: Record Customer)
//     begin
//         // SetPrinter()
//         ws.PageSetup.Orientation := 1; //Portrait
//         //ws.PageSetup.FitToPagesWide := 1;
//         //ws.PageSetup.FitToPagesTall := 200;
//         ws.PageSetup.CenterFooter := 'Seite &S';
//         ws.PageSetup.LeftHeader := Deb.Name;

//         ws.PageSetup.LeftMargin := xlApp.InchesToPoints(0.27559);
//         ws.PageSetup.RightMargin := xlApp.InchesToPoints(0.27559);
//         ws.PageSetup.TopMargin := xlApp.InchesToPoints(0.826771653);
//         ws.PageSetup.BottomMargin := xlApp.InchesToPoints(0.47244);
//         ws.PageSetup.HeaderMargin := xlApp.InchesToPoints(0.51181);
//         ws.PageSetup.FooterMargin := xlApp.InchesToPoints(0.23622);

//         ws.PageSetup.FirstPageNumber := -4105; //xlAutomatic
//         ws.PageSetup.Zoom := 100;
//     end;

//     //[Scope('OnPrem')]
//     procedure AddLine(Beginx: Decimal; Beginy: Decimal; Endx: Decimal; Endy: Decimal)
//     var
//         vRot: Integer;
//     begin
//         // AddLine()
//         vRot := 10;
//         xlLine := ws.Shapes.AddLine(Beginx, Beginy, Endx, Endy);
//         xlLineFormat := xlLine.Line;
//         xlLineFormat.Weight(0.5);
//         xlLineFormat.DashStyle(1);  // msoLineSolid
//         xlLineFormat.Style(1);      // msoLineSingle
//         xlLineFormat.Visible(-1);   // msoTrue
//         xlColor := xlLineFormat.ForeColor;

//         //xlColor.SchemeColor:=10;  //  rot
//         xlLineFormat.BeginArrowheadStyle(1);  //  msoArrowheadNone
//         xlLineFormat.EndArrowheadStyle(1);    //  msoArrowheadNone
//     end;

//     //[Scope('OnPrem')]
//     procedure GetRangeWidth(x1: Integer; y1: Integer; x2: Integer; y2: Integer): Decimal
//     begin
//         // GetRangeWidth
//         rg := ws.Range(ColLetter(x1) + Format(y1) + ':' + ColLetter(x2) + Format(y2));
//         exit(rg.Width);
//     end;

//     //[Scope('OnPrem')]
//     procedure GetRangeHeight(x1: Integer; y1: Integer; x2: Integer; y2: Integer): Decimal
//     begin
//         // GetRangeHeight
//         rg := ws.Range(ColLetter(x1) + Format(y1) + ':' + ColLetter(x2) + Format(y2));

//         exit(rg.Height);
//     end;

//     //[Scope('OnPrem')]
//     procedure NewPage(row: Integer)
//     begin
//         xlApp.ActiveWindow.View := 2; //xlPagebreakView
//         ws.HPageBreaks.Add((ws.Range(Format(row) + ':' + Format(row))));
//     end;

//     //[Scope('OnPrem')]
//     // procedure AddBemerkung(ptop: Decimal; pProjektauslastung: Record PA_Jobs)
//     // begin
//     // end;

//     //[Scope('OnPrem')]
//     procedure SetFontSize(pSize: Integer)
//     begin
//         rg := ws.Cells;
//         rg.Font.Size := pSize;
//     end;

//     //[Scope('OnPrem')]
//     procedure setCellUnderline(col: Integer; row: Integer)
//     begin
//         ws.Range(ColLetter(col) + Format(row)).Font.Underline := 2;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetVerteiler(pVerteiler: Text[250])
//     begin
//         txtVerteiler := pVerteiler;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetRowautofit(row: Integer)
//     begin
//         rg := ws.Rows.Item(row);
//         rg.AutoFit;
//     end;

//     //[Scope('OnPrem')]
//     procedure InsertLogo(x: Integer)
//     var
//         Firma: Record "Company Information";
//         PicPath: Text[250];
//         LocalPicPath: Text[250];
//     begin
//         // Logo
//         PicPath := 'c:\temp\schwarz.bmp';
//         LocalPicPath := FileMgmt2.getTempPath + 'schwarz.bmp';
//         Firma.Get;
//         Firma.CalcFields(Firma.Picture);
//         if Firma.Picture.HasValue then begin
//             Firma.Picture.Export(PicPath);
//             Filemgmt.DownloadToFile(PicPath, LocalPicPath);
//             GetPicture(x, 1, LocalPicPath, 100, 30);
//         end;
//     end;

//     //[Scope('OnPrem')]
//     procedure RunProcedure(strProcedureName: Text[250])
//     begin
//         xlApp.Run(strProcedureName);
//     end;

//     //[Scope('OnPrem')]
//     procedure InsertNewRow(row: Integer)
//     var
//         rg: dotnet xlRange;
//         xlDown: Integer;
//     begin
//         xlDown := -4121;
//         ws.Range(Format(row) + ':' + Format(row)).Insert(xlDown);
//     end;

//     //[Scope('OnPrem')]
//     procedure DisplayAlerts(vbool: Boolean)
//     begin
//         xlApp.DisplayAlerts := vbool;
//     end;

//     //[Scope('OnPrem')]
//     procedure PrintOut2()
//     begin
//         ws.PrintOut();
//     end;

//     //[Scope('OnPrem')]
//     procedure PrintOut()
//     begin
//         ws._PrintOut();
//     end;

//     //[Scope('OnPrem')]
//     procedure SetGroup(FirstRow: Integer; LastRow: Integer)
//     begin
//         ws.Range(Format(FirstRow) + ':' + Format(LastRow)).Group;
//     end;

//     //[Scope('OnPrem')]
//     procedure CopyRange(col: Integer; row: Integer; col2: Integer; row2: Integer; col3: Integer; row3: Integer)
//     begin

//         rg := ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(col2) + Format(row2));
//         rg2 := ws.Range(ColLetter(col3) + Format(row3));
//         rg.Copy(rg2);
//     end;

//     //[Scope('OnPrem')]
//     procedure InsertLine(vzeile: Integer)
//     var
//         xlDown: Integer;
//     begin
//         xlDown := -4121;
//         rg := ws.Range(Format(vzeile) + ':' + Format(vzeile));
//         rg.Insert(xlDown);
//     end;

//     //[Scope('OnPrem')]
//     procedure InsertCol(vCol: Integer)
//     var
//         xlRight: Integer;
//     begin
//         xlRight := -4161;  //Spalte einfügen
//         rg := ws.Range(Format(ColLetter(vCol)) + ':' + Format(ColLetter(vCol)));
//         rg.Insert(xlRight);
//     end;

//     //[Scope('OnPrem')]
//     procedure TurnCell(col: Integer; row: Integer)
//     begin
//         ws.Range(ColLetter(col) + Format(row)).Orientation := 90;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetScreenUpdating(boolIsUpdating: Boolean)
//     begin
//         xlApp.ScreenUpdating(boolIsUpdating);
//     end;

//     //[Scope('OnPrem')]
//     procedure GetMaxWorksheets() ret: Integer
//     begin
//         ret := wb.Sheets.Count();
//     end;

//     //[Scope('OnPrem')]
//     procedure SetBlattschutz(boolProtect: Boolean)
//     begin
//         if boolProtect then
//             ws.Protect
//         else
//             ws.Unprotect;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetCellMultiAlignment(col: Integer; row: Integer; xcol: Integer; xrow: Integer; position: Integer)
//     var
//         pos: Integer;
//     begin
//         //-4108 center
//         //-4152 xlRight
//         case position of
//             0:
//                 pos := -4131;  // xlleft
//             1:
//                 pos := -4108;  // xlcenter
//             2:
//                 pos := -4152;  // xlright
//         end;

//         ws.Range(ColLetter(col) + Format(row) + ':' + ColLetter(xcol) + Format(xrow)).HorizontalAlignment := pos;
//     end;

//     //[Scope('OnPrem')]
//     procedure CreateChart()
//     begin
//         AktChart := wb.Charts.Add;
//         AktChart.Name := 'Liniendiagram TEST';
//         AktChart.ChartType := 65;
//         AktChart.Location(2, 'Kennzahlen'); //Chart as Object
//         rg := xlApp.Range('Kennzahlen!$A$1:$F$3');
//         AktChart.SetSourceData(rg);
//     end;

//     //[Scope('OnPrem')]
//     procedure ExportAsPDF(Pfad: Text)
//     begin
//         wb.ExportAsFixedFormat(0, Pfad);
//     end;

//     //[Scope('OnPrem')]
//     procedure CopyWorksheetBeforeLast(index: Integer)
//     var
//         wss: Dotnet xlSheets;
//         iwss: Integer;
//     begin
//         ws := wb.Worksheets.Item(index);
//         wss := wb.Worksheets;
//         iwss := wb.Worksheets.Count;
//         ws.Copy(wss.Item(iwss));
//         ws.Activate;
//     end;

//     //[Scope('OnPrem')]
//     procedure CopyWorksheet(index: Integer)
//     begin
//         ws := wb.Worksheets.Item(index);
//         ws.Copy(wb.Worksheets.Item(index + 1));
//         ws := wb.Worksheets.Item(index + 1);
//         ws.Activate;
//     end;

//     //[Scope('OnPrem')]
//     procedure GetWorksheetNameByIndex(Index: Integer) ret: Text
//     var
//         wstemp: dotnet xlWorksheet;
//     begin
//         wstemp := wb.Worksheets.Item(Index);
//         ret := wstemp.Name;
//     end;

//     //[Scope('OnPrem')]
//     procedure SetPrinterByName(Printername: Text)
//     var
//         i: Integer;
//     begin
//         //xlApp.ActivePrinter:=Printername;
//         for i := 1 to wb.Worksheets.Count do begin
//             ws := wb.Worksheets.Item(i);
//             ws.Activate;
//             ws._PrintOut(1, 1, 1, false, Printername, false, false);
//         end;
//     end;

//     //[Scope('OnPrem')]
//     procedure RunProcedureParam(strProcedureName: Text[250]; var param: InStream)
//     var
//         [RunOnClient]
//         [SuppressDispose]
//         image: DotNet Bitmap;
//         [RunOnClient]
//         [SuppressDispose]
//         image2: DotNet Image;

//     begin
//         image := image.Bitmap(param);
//         Message(image.ToString());
//         xlApp.Run(strProcedureName, image);
//     end;

//     //[Scope('OnPrem')]
//     procedure RunProcedureParam2(strProcedureName: Text[250]; param: DotNet String)
//     begin
//         xlApp.Run(strProcedureName, param.ToString());
//     end;

//     //[Scope('OnPrem')]
//     procedure SetBlobToPicture(var TempBlob: codeunit "Temp Blob"; x: Decimal; y: Integer; width: Decimal; height: Decimal)
//     var
//         instr: InStream;
//         outstr: OutStream;
//         f: File;
//         filename: Text[250];
//         LocalPicPath: Text[250];

//     begin

//         filename := 'pic.png';

//         if TempBlob.HasValue then begin
//             TempBlob.CreateInStream(instr);
//             f.Create('c:\temp\' + filename);
//             f.CreateOutStream(outstr);
//             CopyStream(outstr, instr);
//             f.Close;

//             LocalPicPath := FileMgmt2.getTempPath + filename;
//             Filemgmt.DownloadToFile('c:\temp\' + filename, LocalPicPath);
//             GetPicture(x, y, LocalPicPath, width, height);
//         end;
//         //MemoryStream:=MemoryStream.MemoryStream();
//         //COPYSTREAM(MemoryStream,ins);
//         //Bytes:=MemoryStream.GetBuffer();
//         //image:=image.Object();
//         //MESSAGE(Convert.ToBase64String(Bytes));
//     end;
// }


