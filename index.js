window.onload = () => {
  const toComma = (val) =>
    Number(val).toLocaleString("ko-KR", { maximumSignificantDigits: 10 });

  document.getElementById("input_excel").addEventListener("change", (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const sheetJson = XLSX.utils.sheet_to_json(workbook.Sheets.sheet1, {
        header: 1,
      });

      const formatJson = sheetJson.reduce((acc, curr, idx) => {
        if (curr.length < 2) {
          return acc;
        }

        const date = curr[1]?.slice(0, 7);
        const type = ["업무추진비(팀비)", "회식비"].includes(curr[6]);

        if (!type) {
          return acc;
        }

        if (!acc) {
          acc = {};
        }

        if (!acc[date] || !acc[date][curr[6]]) {
          acc[date] = {
            ...(acc?.[date] && acc?.[date]),
            [curr[6]]: 0,
          };
        }

        acc[date][curr[6]] += curr[5];

        return acc;
      }, {});

      console.log(JSON.stringify(formatJson));

      const member = document.getElementById("input_member").value;
      const tableData = Object.keys(formatJson).map(
        (date) => `
      <tr>
        <td colspan="2">${date}</td>
      </tr>
      <tr>
        <td>팀비 지급액</td> 
        <td>${toComma(50000 * member)}</td> 
      </tr>
      <tr>
        <td>팀비 사용액</td> 
        <td>${toComma(formatJson[date]["업무추진비(팀비)"])}</td> 
      </tr>
      <tr>
        <td>팀비 차액</td> 
        <td>${toComma(
          50000 * member - formatJson[date]["업무추진비(팀비)"]
        )}</td> 
      </tr>
      `
      );

      document.getElementById("result").innerHTML = tableData.join("");
    };

    reader.readAsArrayBuffer(file);
  });
};
