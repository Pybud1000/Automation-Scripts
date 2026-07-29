function EmbedAllVideos() {
    var ws = ActiveSheet;
    var folderPath = "C:\\Users\\PCXPC\\Documents\\REPORTS\\MONTHLY\\AS COLLECTIONS\\VIDEOS\\ENCHEN\\";
    
    // Find the last row (with data) from the source column
    var lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row;
    
    var embeddedCount = 0;
    var notFoundCount = 0;
    var failedCount = 0;
    
    // Loop through each row
    for (var i = 1; i <= lastRow; i++) {
        var orderID = ws.Range("A" + i).Value2;
        
        if (orderID && orderID.toString().trim() !== "") {
            orderID = orderID.toString().trim();
            var videoPath = folderPath + orderID + ".mp4";
            
            // Attempt Video embedding
            try {
                // Find the target cell, aligned with the row of the Order ID
                var targetCell = ws.Range("B" + i);
                targetCell.Select();
                
                // Delete existing data in the target cell
                var shapes = ws.Shapes;
                for (var s = shapes.Count; s >= 1; s--) {
                    var shp = shapes.Item(s);
                    if (shp.TopLeftCell.Row === i && shp.TopLeftCell.Column === 2) {
                        shp.Delete();
                    }
                }
                
                // Use WPS's AddOLEObject method 
                var oleObj = ActiveSheet.Shapes.AddOLEObject(
                    "",                          // ClassType (empty for file)
                    videoPath,                   // Filename
                    false,                       // Link (false = embed)
                    undefined,                   // (optional)
                    "%SystemRoot%\\system32\\wmploc.dll",  // Icon file
                    -730,                        // Icon index
                    orderID + ".mp4",            // Display name
                    targetCell.Left + 2,         // Left
                    targetCell.Top + 2,          // Top
                    targetCell.Width - 4,        // Width
                    targetCell.Height - 4        // Height
                );
                
                if (oleObj) {
                    ws.Range("C" + i).Value2 = "✓ Embedded";
                    ws.Range("C" + i).Font.Color = 65280; // Green
                    embeddedCount++;
                } else {
                    ws.Range("C" + i).Value2 = "✗ Failed";
                    ws.Range("C" + i).Font.Color = 255; // Red
                    failedCount++;
                }
            } catch (err) {
                ws.Range("C" + i).Value2 = "✗ Error: " + err.message;
                ws.Range("C" + i).Font.Color = 255;
                failedCount++;
            }
        }
    }
    
    // Auto-fit columns
    ws.Columns("A:C").AutoFit();
    
    // Show summary
    alert("Video Embedding Complete!\n\n" +
          "Embedded: " + embeddedCount + " videos\n" +
          "Not found: " + notFoundCount + " videos\n" +
          "Failed: " + failedCount + " videos\n\n" +
          "Check column C for status.");
}
