// Import modules
const axios = require('axios');
const xlsx = require('xlsx');
const fs = require('fs');
const ExcelJS = require('exceljs');
const querystring = require('querystring');


let dataFromXlsx = []
let newData = []
function xoaKiTuTruocDauCach(chuoi) {
  const viTriDauCach = chuoi.indexOf(' ');
  if (viTriDauCach !== -1) {
    return chuoi.slice(viTriDauCach + 1);
  }
  return chuoi;
}
function xoaKiTuTruoc(chuoi) {
  const viTriDauCach = chuoi.indexOf('-');
  if (viTriDauCach !== -1) {
    return chuoi.slice(viTriDauCach + 1);
  }
  return chuoi;
}
function readXlsxFile(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
  let raw_data = []
  data.forEach(item => {
    raw_data.push({
      id: item[0],
      new_ID: xoaKiTuTruocDauCach(item[0].toString()),
    })
  })
  return raw_data;
}

async function readFile() {
  const filePath = "raw.xlsx"; // Replace with your XLSX file path
  dataFromXlsx = readXlsxFile(filePath);
  // console.log(dataFromXlsx)
  // console.log(xoaKiTuTruocDauCach("84 338891404"))
}

async function getInfo(obj) {
  const apiUrl = 'https://dynamic.classin.com/saasajax/teacher.ajax.php?action=getSchoolTeacherFullList';
  const headers = {
    "cookie": "_eeos_uid=7680090;  _eeos_domain=classin.com; _eeos_traffic=fXPpWQCmPZsv%2BTIhWyKUGXWTQ7mWI0ATkyt8JEP6%2FHFFKVMEOdNTkUV6oj1E%2BYV609tVk4Bs8iY%3D"
  };
  const body = {
    page: 1,
    perpage: 20,
    skey: parseInt(obj.new_ID),
    status: 0
  }
  try {
    const response = await axios.post(apiUrl, querystring.stringify(body), { headers });
    if (response.data.error_info.errno == 1 && response.data.data.list.length >0) {
      let a = {
        id: obj.new_ID,
        acc: xoaKiTuTruoc(response.data.data.list[0].account),
        uid: response.data.data.list[0].uid,
        name: response.data.data.list[0].name,
      }
      return a
    }else{
      let a = {
        id: obj.new_ID,
        acc: "ERROR",
        uid: "ERROR",
        name: "ERROR",
      }
      return a
    }
  }
  catch (error) {
    console.error(`Lỗi khi gọi API cho ID (${obj.name}):`, error);
    // Ném lỗi để quản lý ngoại lệ
    throw error;
  }
}
getInfo({
  id:"737743291",
  new_ID:"737743291"
})
// getInfo({
//   id: "933305065",
//   new_ID: "933305065"
// })
readFile()
async function runSequentially() {
  const updatedList = [];
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Data');
  const header = ['id', 'acc', 'uid', "name"];
  worksheet.addRow(header);
  for (const [i, obj] of dataFromXlsx.entries()) {
    try {
      const updatedObj = await getInfo(obj);
      updatedList.push(updatedObj);
      console.log(i,updatedObj)
      worksheet.addRow([updatedObj.id, updatedObj.acc, updatedObj.uid, updatedObj.name])
    } catch (error) {
      // worksheet.addRow([updatedObj.id, updatedObj.acc, 0, 0])
      // Xử lý lỗi tại đây nếu cần
    }
  }
  workbook.xlsx.writeFile('data.xlsx')
    .then(() => {
      console.log('File xlsx đã được tạo!');
    })
    .catch(err => {
      console.error('Lỗi khi tạo file xlsx:', err);
    });
  // Xử lý danh sách đã cập nhật ở đây nếu cần
  console.log('Danh sách đã cập nhật:', updatedList);
}
runSequentially();