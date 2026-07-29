function EmbedAllVideos() {
    var ws = ActiveSheet;
    var folderPath = "";
    
      // Find the last row (with data) from the source column
    var lastRow = ws.Cells(ws.Rows.Count, 1).End(-4162).Row;
    
    var embeddedCount = 0;
    var notFoundCount = 0;
    var failedCount = 0;
    
    var margin = 6;  // 6px margin on each side
    
    // Loop through each row
    for (var i = 1; i <= lastRow; i++) {    
        var orderID = ws.Range("A" + i).Value2;
        
        if (orderID && orderID.toString().trim() !== "") {
            orderID = orderID.toString().trim();
            var videoPath = folderPath + orderID + ".mp4";
            
            // Attempt Video embedding
            try {
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
                
                // Use cell-based sizing with margin
                var oleObj = ActiveSheet.Shapes.AddOLEObject(
                    "",
                    videoPath,
                    false,
                    undefined,
                    "%SystemRoot%\\system32\\wmploc.dll",
                    -730,
                    orderID + ".mp4",
                    targetCell.Left + margin,           // Left with margin
                    targetCell.Top + margin,            // Top with margin
                    targetCell.Width - (margin * 2),    // Width minus margins
                    targetCell.Height - (margin * 2)    // Height minus margins
                );
                
                if (oleObj) {
                    oleObj.Placement = 1;
                    ws.Range("C" + i).Value2 = "✓ Embedded";
                    ws.Range("C" + i).Font.Color = 65280;
                    embeddedCount++;
                } else {
                    ws.Range("C" + i).Value2 = "✗ Failed";
                    ws.Range("C" + i).Font.Color = 255;
                    failedCount++;
                }
            } catch (err) {
                ws.Range("C" + i).Value2 = "✗ Error: " + err.message;
                ws.Range("C" + i).Font.Color = 255;
                failedCount++;
            }
        }
    }
    
    ws.Columns("A:C").AutoFit();
    alert("Video Embedding Complete!\n\n" +
          "✓ Embedded: " + embeddedCount + " videos\n" +
          "✗ Not found: " + notFoundCount + " videos\n" +
          "⚠ Failed: " + failedCount + " videos\n\n" +
          "Check column C for status.");
}
