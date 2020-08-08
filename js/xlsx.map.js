(function () {
    var changeColumns = {
        B1: "Mã Kiện Hàng",
        D1: "Trạng Thái Đơn Hàng",
        G1: "Đơn Vị Vận Chuyển",
        V1: "Người bán trợ giá",
        X1: "Tổng số tiền được người bán trợ giá",
        AA1: "Tổng giá bán (sản phẩm)",
        AB1: "Tổng giá trị đơn hàng (VND)",
        AE1: "Mã giảm giá của Shopee",
        AM1: "Tổng số tiền người mua thanh toán",
    }
    var fileCatcher = document.getElementById("file-catcher");
    var fileInput = document.getElementById("file-input");
    var fileListDisplay = document.getElementById("file-list-display");
    var msg = document.getElementsByTagName("h1");
    var smbtn = document.getElementById("submit-btn");

    var fileList = [];
    var renderFileList, sendFile;

    fileCatcher.addEventListener("submit", function (evnt) {
        evnt.preventDefault();
        // fileList.forEach(async function (file) {
        //     if (count == 9) {
        //         await new Promise((r) => setTimeout(r, 2000));
        //         count = 0;
        //     }
        //     XLSX.writeFile(
        //         file.data,
        //         file.name.name
        //     );
        //     count += 1;
        // });
    });

    fileInput.addEventListener("change", async function (evnt) {
        fileList = [];
        var count = 0;
        for (var i = 0; i < fileInput.files.length; i++) {
            if (count == 8) {
                await new Promise((r) => setTimeout(r, 2000));
                count = 0;
            }
            handleFile(evnt, i);

            fileList.push({
                name: fileInput.files[i],
                // data: file
            });
            count += 1
        }
        renderFileList();
        msg[0].innerHTML = '<span class="badge badge-pill badge-warning">Rendered</span>';
        // smbtn.disabled = false
    });

    renderFileList = function () {
        fileListDisplay.innerHTML = "";
        fileList.forEach(function (file, index) {
            var fileDisplayEl = document.createElement("p");
            // fileDisplayEl.innerHTML = index + 1 + ": " + file.name.name;
            fileDisplayEl.innerHTML = `<div class="uploaded uploaded">
            <i class="far fa-file-excel"></i>
            <div class="file">
                <div class="file__name">
                    <p>${file.name.name}</p>
                    <p>${index+1}</p>
                </div>
                <div class="progress">
                    <div class="progress-bar bg-success progress-bar-striped progress-bar-animated"
                        style="width:100%"></div>
                </div>
            </div>
        </div>`

            fileListDisplay.appendChild(fileDisplayEl);
        });
    };

    async function handleFile(e, inx = 0) {
        var files = e.target.files,
            f = files[inx];

        // console.log(f);
        var reader = new FileReader();
        reader.onload = async function (e) {
            var data = new Uint8Array(e.target.result);
            var workbook = XLSX.read(data, {
                type: "array"
            });
            console.log(f.name, workbook);


            let sheet = workbook.Sheets[workbook.SheetNames[0]];
            for (const column in changeColumns) {
                // console.log(sheet[column]);
                sheet[column].v = changeColumns[column]
                sheet[column].w = changeColumns[column]
                // console.log(sheet[column]);
            }

            // console.log(sheet.AA1);

            XLSX.writeFile(workbook, f.name);
            // return
        };
        reader.readAsArrayBuffer(f);
    }
    // fileInput.addEventListener("change", handleFile, false);
})();