const { ipcRenderer } = require("electron");
const admZip = require("adm-zip");
const JSZip = require("jszip"); // 引入 JSZip
const docx4js = require("docx4js");

const filePathInput = document.getElementById("fileInput"); // 假设你有一个输入框用于输入文件路径
const outputDiv = document.getElementById("output");

filePathInput.addEventListener("click", (event) => {
  event.preventDefault(); // 阻止默认行为
  ipcRenderer.send("open-file-dialog"); // 发送消息到主进程
});

ipcRenderer.on("selected-file", async (event, filePath) => {
  // 处理接收到的文件路径
  console.log("选择的 zip 文件路径：", filePath);

  try {
    //  直接使用文件路径创建 adm-zip 实例
    const zip = new admZip(filePath);
    const zipEntries = zip.getEntries();

    outputDiv.innerText = "正在处理文件中...";

    const zipWriter = new JSZip();

    for (const entry of zipEntries) {
      if (entry.entryName.endsWith(".zip")) {
        const innerZip = new admZip(entry.getData());
        const innerZipEntries = innerZip.getEntries();

        for (const innerEntry of innerZipEntries) {
          if (
            innerEntry.entryName.endsWith(".docx") ||
            innerEntry.entryName.endsWith(".doc")
          ) {
            let fileData = innerEntry.getData();
            // 如果是 .doc 文件，则转换为 .docx
            // if (innerEntry.entryName.endsWith(".doc")) {
            //   const doc = await docx4js.load(fileData);
            //   const docx = new docx4js();
            //   docx.attachModule(new docx4js.SaveModule());
            //   docx.officeDocument = doc.officeDocument;
            //   fileData = await docx.saveAsBlob("docx");
            // }
            // 将文件添加到 zipWriter 中
            zipWriter.file(innerEntry.entryName, fileData); // 将 .doc 后缀替换为 .docx
          }
        }
      }
    }

    zipWriter.generateAsync({ type: "blob" }).then((content) => {
      const downloadLink = document.createElement("a");
      downloadLink.style.display = "block";
      downloadLink.style.width = "200px";
      downloadLink.style.height = "30px";
      const url = URL.createObjectURL(content);
      downloadLink.href = url;
      downloadLink.download = "extracted_docs.zip";
      downloadLink.textContent = "下载提取后的文件";
      outputDiv.innerHTML = `<p>提取完成！</p><p><a href="#" id="downloadLink">下载提取后的文件</a></p>`;
      document.getElementById("downloadLink").replaceWith(downloadLink);
    });
  } catch (error) {
    console.error("读取或处理 zip 文件时出错:", error);
    outputDiv.innerText = "处理文件时出错！";
  }
});
