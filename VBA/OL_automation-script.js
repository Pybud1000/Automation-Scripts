function EmbedAllVideos() {
    var ws = ActiveSheet;
    var folderPath = "C:\\Users\\PCXPC\\Documents\\REPORTS\\VIDEO DUMP\\";
    
    // Find the last row (with data) from the source column
    var lastRow = ws.Cells(ws.Rows.Count, 4).End(-4162).Row;
    
    var embeddedCount = 0;
    var notFoundCount = 0;
    var failedCount = 0;
    
    var margin = 6;  // 6px margin on each side
    
    // Loop through each row
    for (var i = 1; i <= lastRow; i++) {    
        var orderID = ws.Range("D" + i).Value2;  // Source: Column D
        
        if (orderID && orderID.toString().trim() !== "") {
            orderID = orderID.toString().trim();
            var videoPath = folderPath + orderID + ".mp4";
            
            var fileExists = false;
            try {
                // WPS has a built-in file check using the FileSystemObject
                // but if it's not available, we'll just catch and continue
                var fso = new ActiveXObject("Scripting.FileSystemObject");
                fileExists = fso.FileExists(videoPath);
            } catch (e) {
                // If ActiveXObject fails, try WPS's built-in method
                try {
                    // WPS uses a different object for file checking
                    fileExists = (Dir(videoPath) !== "");
                } catch (e2) {
                    // If all else fails, assume the file might exist and try to embed
                    fileExists = true;
                }
            }
            
            // If file doesn't exist, mark and skip WITHOUT calling AddOLEObject
            if (!fileExists) {
                ws.Range("O" + i).Value2 = "✗ File not found";
                ws.Range("O" + i).Font.Color = 49407; // Orange
                notFoundCount++;
                continue;  // Skip to next row - NO POPUP
            }
            
            // Attempt Video embedding - only if file exists
            try {
                var targetCell = ws.Range("N" + i);  // Target: Column N
                targetCell.Select();
                
                // Delete existing data in the target cell
                var shapes = ws.Shapes;
                for (var s = shapes.Count; s >= 1; s--) {
                    var shp = shapes.Item(s);
                    if (shp.TopLeftCell.Row === i && shp.TopLeftCell.Column === 14) {
                        shp.Delete();
                    }
                }
                
                // Embed the video
                var oleObj = ActiveSheet.Shapes.AddOLEObject(
                    "",
                    videoPath,
                    false,
                    undefined,
                    "%SystemRoot%\\system32\\wmploc.dll",
                    -730,
                    orderID + ".mp4",
                    targetCell.Left + margin,
                    targetCell.Top + margin,
                    targetCell.Width - (margin * 2),
                    targetCell.Height - (margin * 2)
                );
                
                if (oleObj) {
                    oleObj.Placement = 1;
                    ws.Range("O" + i).Value2 = "✓ Embedded";
                    ws.Range("O" + i).Font.Color = 65280;
                    embeddedCount++;
                } else {
                    ws.Range("O" + i).Value2 = "✗ Failed";
                    ws.Range("O" + i).Font.Color = 255;
                    failedCount++;
                }
            } catch (err) {
                ws.Range("O" + i).Value2 = "✗ Error";
                ws.Range("O" + i).Font.Color = 255;
                failedCount++;
            }
        }
    }
    
    // FIX 1: Auto-fit each column separately (was causing error)
    ws.Columns("D").AutoFit();
    ws.Columns("N").AutoFit();
    ws.Columns("O").AutoFit();
    
    // FIX 2: Use status bar instead of alert (was causing error)
    Application.StatusBar = "✓ Embedded: " + embeddedCount + 
                            " | ✗ Not found: " + notFoundCount + 
                            " | ⚠ Failed: " + failedCount;
}