const info = document.querySelector("#info");
// info.innerText = `Current Chrome (v${window.versions.chrome()})`;

const req = async () => {
  // const resp = await window.versions.ping();
  // console.log(resp);
};
req();

const importer = {
  docInput: null,
  xlsInput: null,
  docButton: null,
  xlsButton: null,
  downloadButton: null,
  tips: null,
  container: null,
  init() {
    this.container = document.querySelector(".main-container");
    this.downloadButton = document.querySelector(".download-area");
    this.tips = document.querySelector(".tips");
    this.docButton = document.querySelector(".doc");
    this.docInput = document.getElementById("doc");
    this.xlsButton = document.querySelector(".xls");
    this.xlsInput = document.getElementById("xls");
    this.arrow = document.querySelector(".to");
    this.docButton.addEventListener("click", importer.clickDoc);
    this.xlsButton.addEventListener("click", importer.clickXls);
    this.downloadButton.addEventListener("click", importer.clickDownload);
    this.tips = document.querySelector(".tips");
    // this.docInput.addEventListener("change", async (e) => {
    //   const file = e.target.files[0];
    //   document.getElementById("content").innerHTML = "<p>正在解析文档...</p>";
    //   const buffer = await file.arrayBuffer();
    //   const result = await mammoth.convertToHtml({ buffer });
    //   // const result = await window.electronAPI.parseWordFile(buffer);
    //   console.log(result);
    //   document.getElementById("content").innerHTML = result.value;
    // });
    console.log("[importer]init completed.");
    this.show();
  },
  show() {
    this.container.classList.remove("hajime");
    this.container.classList.add("show");
    setTimeout(() => {
      this.container.classList.remove("show");
    }, 2000);
  },
  switchTip() {
    this.tips.classList.add("active");
    setTimeout(() => {
      const firstTip = document.querySelector(".tip:nth-child(1)");
      const parent = firstTip.parentElement;
      parent.classList.remove("active");
      parent.removeChild(firstTip);
      parent.appendChild(firstTip);
    }, 1000);
  },
  getTime() {
    let date = new Date();
    let el = document.createElement("p");
    let hour = date.getHours() < 10 ? "0" + date.getHours() : date.getHours();
    let minute =
      date.getMinutes() < 10 ? "0" + date.getMinutes() : date.getMinutes();
    let second =
      date.getSeconds() < 10 ? "0" + date.getSeconds() : date.getSeconds();
    return `[${hour}:${minute}:${second}]`;
  },
  async clickDownload() {
    const res = await window.electronAPI.saveFile();
    if (res.success) {
      importer.downloadButton.classList.remove("focus");
      importer.switchTip();
      let sel = document.createElement("p");
      sel.innerHTML = `${importer.getTime()}  保存成功`;
      document.getElementById("content").appendChild(sel);
      importer.downloadButton.classList.remove("done");
      importer.docButton.classList.remove("done");
      importer.xlsButton.classList.remove("done");
    } else {
      let sel = document.createElement("p");
      sel.innerHTML = `${importer.getTime()}  保存失败: ${res.error}`;
      document.getElementById("content").appendChild(sel);
    }
  },
  async clickDoc() {
    // importer.docInput.click();
    const filePath = await window.electronAPI.openFileDialog(["docx"]);
    if (!filePath) return;
    let el = document.createElement("p");
    el.innerHTML = `${importer.getTime()}  正在加载文档...`;
    document.getElementById("content").appendChild(el);
    const result = await window.electronAPI.parseWordFile(filePath);
    if (result.success) {
      let sel = document.createElement("p");
      sel.innerHTML = `${importer.getTime()}  完成解析Word数据`;
      document.getElementById("content").appendChild(sel);
      importer.docButton.classList.remove("focus");
      importer.docButton.classList.add("done");
      importer.arrow.classList.add("process");
      setTimeout(() => {
        importer.downloadButton.classList.add("focus");
        importer.downloadButton.classList.add("done");
        importer.arrow.classList.remove("process");

        importer.switchTip();
      }, 8000);
    } else {
      let sel = document.createElement("p");
      sel.innerHTML = `${importer.getTime()}  解析失败: ${result.error}`;
      document.getElementById("content").appendChild(sel);
    }
  },
  async clickXls() {
    const filePath = await window.electronAPI.openFileDialog(["xlsx"]);
    if (!filePath) return;
    let el = document.createElement("p");
    el.innerHTML = `${importer.getTime()}  正在加载文档...`;
    document.getElementById("content").appendChild(el);
    const result = await window.electronAPI.parseExcelFile(filePath);
    if (result.success) {
      let sel = document.createElement("p");
      sel.innerHTML = `${importer.getTime()}  已加载${result.data.length}条数据`;
      document.getElementById("content").appendChild(sel);
      importer.xlsButton.classList.remove("focus");
      importer.docButton.classList.add("focus");
      importer.xlsButton.classList.add("done");
      importer.switchTip();
    } else {
      let sel = document.createElement("p");
      sel.innerHTML = `${importer.getTime()}  解析失败: ${result.error}`;
      document.getElementById("content").appendChild(sel);
    }
  },
};

window.onload = () => {
  importer.init();
};
