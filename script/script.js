
var JSONData = [
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
$('.excel').click(function(){
    JSONToCSVConvertor();
});

function JSONToCSVConvertor(JSONData) {
    const resultclone = $('.creatExcelTable')
    resultclone.children('tbody').children('tr').remove();

      const tr = $('<tr/>');
        $('table').append(tr);
        tr.append(`<th>ldapName</th>`);
        tr.append(`<th>Yükleme Yapan Numara</th>`);
        tr.append(`<th>Miktar</th>`);
        tr.append(`<th>Gönderen Kart</th>`);
        tr.append(`<th>Alıcı Kart</th>`);
        tr.append(`<th>İşlem Tarihi</th>`);
        tr.append(`<th>Kampanya Adı</th>`);
        tr.append(`<th>Durum</th>`);
        tr.append(`<th>Açıklama</th>`);
    
        JSONData.map( item =>  {
       
          let status = ''
          if(item.status === "SUCCESS"){
            status = 'Başarılı'
          } else if (item.status === "ERROR"){
            status = 'Başarısız'
          }else if (item.status === "WAITING") {
            status = 'Beklemede'
          }
            
          const tr = $('<tr/>');
          tr.append(`<td>${item.ldapName}</td>`);
          tr.append(`<td>${item.msisdn}</td>`);
          tr.append(`<td>${item.amount}</td>`);
          tr.append(`<td>${item.senderMaskedCard ? item.senderMaskedCard : ''}</td>`);
          tr.append(`<td>${item.receiverMaskedCard ? item.receiverMaskedCard : ''}</td>`);
          tr.append(`<td>${ReportTitle}</td>`);
          tr.append(`<td>${status ? status : ''}</td>`);
          tr.append(`<td>${item.errorMessage ? item.errorMessage : ''}</td>`);
          $('table').append(tr);
      });
      

    var uri = 'data:application/vnd.ms-excel;charset=utf-8;base64,',
        template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">' +
        '<head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet>' +
        '<x:Name>Ark1</x:Name>' +
        '<x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]-->' +
        '<style> table{mso-displayed-decimal-separator:"\,";mso-displayed-thousand-separator:"\.";} td{border:none;font-family: Calibri, sans-serif;}  .number{mso-number-format:"###,##0";}</style>' +
        '<meta name=ProgId content=Excel.Sheet>' +
        '<meta http-equiv="Content-Type" content="charset=UTF-8">'+
        '</head><body>' +
        '<table>'+ resultclone.html() +'</table>'+
        '</body></html>',
    base64 = function (s) { return window.btoa(unescape(encodeURIComponent(s))); };
        getContent(
          uri + base64(template),
          'xls',
          template,
          'application/vnd.ms-excel'
      );
  };


  function  getContent(href, extension, content, MIME) {
    var downloadAttrSupported = document.createElement('a').download !== undefined;
      var a, blobObject, name = 'file';
      if (window.Blob && window.navigator.msSaveOrOpenBlob) {
      blobObject = new Blob([content]);
      window.navigator.msSaveOrOpenBlob(blobObject, name + '.' + extension);
      } else if(downloadAttrSupported) {
      a = document.createElement('a');
      a.href = href;
      a.target      = '_blank';
      a.download    = name + '.' + extension;
      document.body.appendChild(a);
      a.click();
      a.remove();
      }
  }
