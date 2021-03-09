// The controladdin type declares the new add-in.
/// <summary>
/// ControlAddIn SampleAddIn.
/// </summary>
controladdin BC2XL
{
    
    RequestedHeight = 20;
    MinimumHeight = 20;
    MaximumHeight = 20;

    // The Scripts property can reference both external and local scripts.
    Scripts = 'https://cdnjs.cloudflare.com/ajax/libs/knockout/3.4.2/knockout-debug.js',
               './js/FileSaver.js',
              'https://cdnjs.cloudflare.com/ajax/libs/xlsx-populate/1.21.0/xlsx-populate.min.js',
            'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.2.0/exceljs.min.js',
            './js/main.js';
           // 'https://cdnjs.cloudflare.com/ajax/libs/exceljs/4.2.0/exceljs.min.js';
            // 'https://unpkg.com/xlsx/dist/shim.min.js',
            // 'https://unpkg.com/xlsx/dist/xlsx.full.min.js',
            // 'https://unpkg.com/blob.js@1.0.1/Blob.js';
            
    // The StartupScript is a special script that the web client calls once the page is loaded.
    StartupScript = './js/startup.js';

    // Specifies the StyleSheets that are included in the control add-in.
    StyleSheets = './js/skin.css';

    // Specifies the Images that are included in the control add-in.
    //Images = 'image.png';

    // The procedure declarations specify what JavaScript methods could be called from AL.
    // In main.js code, there should be a global function CallJavaScript(i,s,d,c) {Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('CallBack', [i, s, d, c]);}
    
   /// <summary>
   /// createXLSX.
   /// </summary>
   /// <param name="s">text.</param>
   procedure readXLSX(s: text);

    /// <summary>
    /// createXLSX.
    /// </summary>
    /// <param name="d">JsonObject.</param>
    /// <param name="filename">text.</param>
    procedure createXLSX(d: JsonObject; filename: text);
    
    // The event declarations specify what callbacks could be raised from JavaScript by using the webclient API:
    // Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('CallBack', [42, 'some text', 5.8, 'c'])
    
    event GetWorksheetAsJson(d: JsonObject);
}