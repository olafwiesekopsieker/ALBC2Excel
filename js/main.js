var workbook = new ExcelJS.Workbook();

function readXLSX(s, d, filename){
        
    var b; 
    try
    {
       b=atob(s);
    } 
    catch(e)
    {
       console.error(e);
    }
    
    // console.log(byteArray);
  
    try {
       
        workbook.xlsx.load(b).then(function(workbook){
            // console.timeEnd();
            // var result = ''
            // workbook.worksheets.forEach(function (sheet) {
            //   sheet.eachRow(function (row, rowNumber) {
            //     result += row.values + ' | \n'
            //   })
            // }) 
            // console.log(result);
           //Convert worksheet to JSON-Object and back to BC
           var jsonsheet;
           var ws;
           console.log(workbook);
           ws=workbook.worksheets[0];
           console.log(ws);
           
           //jsonsheet = XLSX.utils.sheet_to_json(ws);
           //console.log(jsonsheet);
        //    Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('GetWorksheetAsJson', [ws]);

           // TODO Create Content
                      
           
        //     // save as Blob
        //     workbook.xlsx.writeBuffer().then(function(buffer) {
        //         XlsxPopulate.fromDataAsync(buffer)
        //             .then(function (workbook) {
        //                 console.log(workbook);
        //                 // save file
        //                 workbook.outputAsync()
        //                 .then(function (blob) {
        //                     if (window.navigator && window.navigator.msSaveOrOpenBlob) {
        //                         // If IE, you must uses a different method.
        //                         window.navigator.msSaveOrOpenBlob(blob, filename);
        //                     } else {
        //                         var url = window.URL.createObjectURL(blob);
        //                         var a = document.createElement("a");
        //                         document.body.appendChild(a);
        //                         a.href = url;
        //                         a.download = filename;
        //                         a.click();
        //                         window.URL.revokeObjectURL(url);
        //                         document.body.removeChild(a);
        //                     }
        //                 });
        //             });
        //     });    
        });

    } catch(error) 
    {
        // Catch compilation errors (errors caused by the compilation of the template : misplaced tags)
        errorHandler(error);
    }

    
   



    //Save Workbook
    // var buff = workbook.xlsx.writeBuffer().then(function (data) {
    //     var blob = new Blob([data], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    //     saveAs(blob, 'fileName');
    // });

    //Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('CallBack', [s]);
    

}



