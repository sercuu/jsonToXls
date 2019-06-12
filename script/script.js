titles = {
  name: {
    tr: 'Ad',
    en: 'Name'
  },
  title: {
    tr: 'Pozisyon',
    en: 'Position'
  },
  department: {
    tr: 'Departman',
    en: 'Department'
  },
  location: {
    tr: 'Lokasyon',
    en: 'Location'
  },
  workingStartDate: {
    tr: 'Başlama Tarihi',
    en: 'Start Date'
  },
  workingEndDate: {
    tr: 'Çıkış Tarihi',
    en: 'Termination Date'
  },
  leaveReason: {
    tr: 'Çıkış Nedeni',
    en: 'Exit Reason'
  }
};

var JSONDatas = [
    {
        identifier: "7e44cf91-177b-4520-93f4-83cd63d6a3e4", 
        amount: 51,
        createdDate: "2019-03-25T10:24:49.66+03:00[Europe/Istanbul]",
        errorMessage: "Ödeme Başarılı",
        fullName: "-",
        identifier: "7e44cf91-177b-4520-93f4-83cd63d6a3e4",
        lastModifiedDate: "2019-03-25T11:00:26.861+03:00[Europe/Istanbul]",
        ldapName: "GLB90039918",
        msisdn: "5078683130",
        receiverMaskedCard: "462276******8324",
        rowVersion: 1,
        senderMaskedCard: "462276******5113",
        status: "SUCCESS"
    },
    {
        identifier: "7e44cf91-177b-4520-93f4-83cd63d6a3e4", 
        amount: 51,
        createdDate: "2019-03-25T10:24:49.66+03:00[Europe/Istanbul]",
        errorMessage: "Ödeme Başarılı",
        fullName: "-",
        identifier: "7e44cf91-177b-4520-93f4-83cd63d6a3e4",
        lastModifiedDate: "2019-03-25T11:00:26.861+03:00[Europe/Istanbul]",
        ldapName: "GLB90039918",
        msisdn: "5078683130",
        receiverMaskedCard: "462276******8324",
        rowVersion: 1,
        senderMaskedCard: "462276******5113",
        status: "SUCCESS"
    },
    {
        identifier: "7e44cf91-177b-4520-93f4-83cd63d6a3e4", 
        amount: 51,
        createdDate: "2019-03-25T10:24:49.66+03:00[Europe/Istanbul]",
        errorMessage: "Ödeme Başarılı",
        fullName: "-",
        identifier: "7e44cf91-177b-4520-93f4-83cd63d6a3e4",
        lastModifiedDate: "2019-03-25T11:00:26.861+03:00[Europe/Istanbul]",
        ldapName: "GLB90039918",
        msisdn: "5078683130",
        receiverMaskedCard: "462276******8324",
        rowVersion: 1,
        senderMaskedCard: "462276******5113",
        status: "SUCCESS"
    }
]
const table = document.getElementById('creatExcelTable');

  const generateTableHead = (tableHTML, title) => {
    const thead = tableHTML.createTHead();
    const row = thead.insertRow();
    for (const key in title) {
      const th = document.createElement('th');
      // you can make language control 
      //const text = document.createTextNode(title[key][language]);
      const text = document.createTextNode(title[key].tr);
      th.appendChild(text);
      row.appendChild(th);
    }
  };

  const generateTable = (tables, JSONData) => {
    JSONData.map(element => {
      const row = tables.insertRow();
      for (const key in element) {
        if (element.hasOwnProperty.call(element, key)) {
          const cell = row.insertCell();
          const text = document.createTextNode(element[key]);
          cell.appendChild(text);
        }
      }
      return '';
    });
  };
  const getContent = (href, extension, content) => {
    const downloadAttrSupported =
      document.createElement('a').download !== undefined;
    let a;
    let blobObject;
    const name = 'file';
    if (window.Blob && window.navigator.msSaveOrOpenBlob) {
      blobObject = new Blob([content]);
      window.navigator.msSaveOrOpenBlob(blobObject, `${name}.${extension}`);
    } else if (downloadAttrSupported) {
      a = document.createElement('a');
      a.href = href;
      a.target = '_blank';
      a.download = `${name}.${extension}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
    }
  };
  const download = () => {
    const uri = 'data:application/vnd.ms-excel;charset=utf-8;base64,';
    const template =
      `${'<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
        '<head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>' +
        '<x:Name>Ark1</x:Name>' +
        '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
        '<style> table{mso-displayed-decimal-separator:",";mso-displayed-thousand-separator:".";} td{border:none;font-family: Calibri, sans-serif;}  .number{mso-number-format:"###,##0";}</style>' +
        '<meta name=ProgId content=Excel.Sheet>' +
        '<meta http-equiv="Content-Type" content="charset=UTF-8">' +
        '</head><body>' +
        '<table>'}${table.innerHTML}</table>` + '</body></html>';
    const base64 = s => {
      return window.btoa(unescape(encodeURIComponent(s)));
    };

    getContent(
      uri + base64(template),
      'xls',
      template,
      'application/vnd.ms-excel'
    );
  };

  $('.excel').click(function(){
    generateTableHead(table, titles);
    generateTable(table, JSONDatas);
    download(table, JSONDatas);
});
