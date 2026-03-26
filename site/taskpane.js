Office.onReady(function(info) {
  if (info.host === Office.HostType.Excel) {
    console.log("Dollinger CAM Tools is ready.");

    const btn = document.getElementById("testBtn");
    if (btn) {
      btn.onclick = () => {
        Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getActiveWorksheet();
          const range = sheet.getRange("A1");
          range.values = [["Hello from Dollinger CAM Tools!"]];
          await context.sync();
        });
      };
    }
  }
});
