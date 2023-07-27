import * as XLSX from "xlsx";
import { useState } from "react";
import * as React from "react";
import "./excel.css";
import SwapHorizIcon from "@mui/icons-material/SwapHoriz";

function Excel() {
  const [jsData, setJsData] = useState([]);
  const [data, setData] = useState([]);
  const [file, setFile] = useState();

  const handleFile1 = (e) => {
    console.log(e.target.files[0]);
    var reader = new FileReader();
    reader.readAsText(e.target.files[0]);
    reader.onload = onReaderLoad;
    function onReaderLoad(event) {
      let obj = JSON.parse(event.target.result);
      console.log(obj);
      setJsData(obj);
      jsData.push(obj);
      setJsData(...jsData);
      setFile(obj);
    }
  };

  const handleFile = async (e) => {
    const file = e.target.files[0];
    const dat = await file.arrayBuffer();
    const book = XLSX.read(dat);
    let shee = book.Sheets[book.SheetNames[0]];
    let sheetData = XLSX.utils.sheet_to_json(shee);
    setData(sheetData);
    data.push(sheetData);
    setData(...data);
    setFile(sheetData);
  };

  let downData = () => {
    if (data.length > 0) {
      const jsonString = `data:abc.json;chatset=utf-8,${encodeURIComponent(
        JSON.stringify(data)
      )}`;
      const link = document.createElement("a");
      link.href = jsonString;
      link.download = "data.json";

      link.click();
      setData([]);
    } else {
      alert("No data to download");
    }
  };
  const downloadData = () => {
    if (jsData.length > 0) {
      const worksheet = XLSX.utils.json_to_sheet(jsData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
      let buffer = XLSX.write(workbook, { bookType: "xlsx", type: "buffer" });
      XLSX.write(workbook, { bookType: "xlsx", type: "binary" });
      XLSX.writeFile(workbook, "data.xlsx");
      // setData(...data,[])
      setJsData([]);
    } else {
      alert("No data to download");
    }
  };
  let toogle = (e) => {
    if (e.type === "click") {
      let h1 = document.getElementsByTagName("h1");
      h1[0].innerHTML = "Json to Excel Convertor";
      let js = document.getElementById("json");
      js.style.display = "block";
      let exc = document.getElementById("excel");
      exc.style.display = "none";
      setFile();
    }
  };
  let toogle1 = (e) => {
    if (e.type === "click") {
      let h1 = document.getElementsByTagName("h1");
      h1[0].innerHTML = "Excel to Json Convertor";
      let js = document.getElementById("json");
      js.style.display = "none";
      let exc = document.getElementById("excel");
      exc.style.display = "block";
      setFile();
    }
  };
  return (
    <div style={{ color: "#21bfdb", paddingTop: "20px" }}>
      <h1 style={{ textAlign: "center" }}>Excel to Json Convertor</h1>
      <div id="json" style={{ display: "none" }}>
        <label class="custom-file-upload">
          <input
            accept="application/json"
            type="file"
            onChange={(e) => handleFile1(e)}
          />
          Upload
        </label>
        <span id="swap-btn" onClick={toogle1}>
          {<SwapHorizIcon />}
        </span>
        <label class="btn" onClick={downloadData}>
          Download
        </label>
      </div>
      <div id="excel">
        <label class="custom-file-upload">
          <input
            accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            type="file"
            onChange={(e) => handleFile(e)}
          />
          Upload
        </label>
        <span id="swap-btn" onClick={toogle}>
          {<SwapHorizIcon />}
        </span>
        <label class="btn" onClick={downData}>
          Download
        </label>
      </div>

      <div id="box" className="box">
        {JSON.stringify(file)}
      </div>
    </div>
  );
}

export default Excel;
