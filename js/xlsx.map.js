(function () {
  var fileCatcher = document.getElementById("file-catcher");
  var fileInput = document.getElementById("file-input");
  var fileListDisplay = document.getElementById("file-list-display");
  var msg = document.getElementsByTagName("h1");

  var fileList = [];
  var renderFileList, sendFile;

  fileCatcher.addEventListener("submit", function (evnt) {
    evnt.preventDefault();
    var count = 0;
    fileList.forEach(async function (file) {
      if (count == 9) {
        await new Promise((r) => setTimeout(r, 2000));
        count = 0;
      }
      saveAs(
        new Blob([file.data], { type: "application/octet-stream" }),
        file.name.name
      );
      count += 1;
    });
  });

  fileInput.addEventListener("change", function (evnt) {
    fileList = [];
    var zip = new JSZip();
    for (var i = 0; i < fileInput.files.length; i++) {
      file = handleFile(evnt, i);
      fileList.push({ name: fileInput.files[i], data: file });
    }
    renderFileList();
    msg[0].innerHTML = '<span style="color: #009688;">Rendered</span>';
  });

  renderFileList = function () {
    fileListDisplay.innerHTML = "";
    fileList.forEach(function (file, index) {
      var fileDisplayEl = document.createElement("p");
      fileDisplayEl.innerHTML = index + 1 + ": " + file.name.name;
      fileListDisplay.appendChild(fileDisplayEl);
    });
  };

  async function handleFile(e, inx = 0, zip) {
    var files = e.target.files,
      f = files[inx];

    console.log(f);
    var reader = new FileReader();
    reader.onload = async function (e) {
      var data = new Uint8Array(e.target.result);
      var workbook = XLSX.read(data, { type: "array" });

      var ws_name = "SheetJS";
      /* make worksheet */
      var ws_data = [
        ["S", "h", "e", "e", "t", "J", "S"],
        [1, 2, 3, 4, 5],
      ];
      var ws = XLSX.utils.aoa_to_sheet(ws_data);
      console.log(f.name, workbook);
      console.log(workbook.Sheets["Các chỉ số quan trọng"].A1.v);
      workbook.Sheets["Các chỉ số quan trọng"].A1.v = "Kiet";
      console.log(workbook.Sheets["Các chỉ số quan trọng"].A1.v);

      /* Add the worksheet to the workbook */
      await XLSX.utils.book_append_sheet(workbook, ws, ws_name);
      console.log(workbook);

      // XLSX.writeFile(workbook, f.name);
      var wopts = {
        bookType: "xlsx",
        bookSST: false,
        type: "binary",
        compression: true,
      };
      return XLSX.write(workbook, wopts);
    };
    reader.readAsArrayBuffer(f);
  }
  // fileInput.addEventListener("change", handleFile, false);
})();