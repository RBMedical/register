window.onload = function() {
    var newURL = "register/?code=" + encodeURIComponent("4/0AVG7fiSARmAs51gJN6VxDl7Y7TtnhkBHe0ZIFIoC5uqVZX_4s0aaNj6C__FoGEri5_O2tA") + 
     "&scope=" + encodeURIComponent("https://www.googleapis.com/auth/spreadsheets");

  window.history.replaceState({ path: newURL }, '', newURL);
}

 function showLoading() {
            document.getElementById('loadingIcon').style.display = 'block';
        }

  function hideLoading() {
            document.getElementById('loadingIcon').style.display = 'none';
        }


// ข้อมูล Client
const apiKey = 'AIzaSyCvtBvaZ5celuXWAZ1aJNuxBsfJEd3nMuA';
const spreadsheetId = '1_aUWV9uDvVn_WBs25ZsHtVLilUYB9iNP87yadjSbHsw';
const rangesheet1 = 'data!A2:ZZ'; 
const rangesheet2 = 'program!A2:ZZ';
const rangesheet3 = 'register!A2:ZZ';
const rangesheet4 = 'register!A2:A';
const rangesheet5 = 'sticker!A2:ZZ';
const rangesheet6 = 'specimencount!A2:ZZ';
const rangesheet7 = 'specimencount!A:A';
const rangesheet8 = 'program!A2:ZZ';
const rangesheet9 = 'sticker!A2:ZZ';
const rangesheet10 = 'register!A2:ZZ';

const clientId = "257006847363-7e1gusb11uc8o4qg3b8u4lg3d9lk8r12.apps.googleusercontent.com";
const clientSecret = "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";
const redirectUri = "https://rbmedical.github.io/register";     
const access_token ='ya29.a0AcM612zIViuQBvE59SU7qoceVTpm6MwDvHQd1eTNE5NUScNZjBhFMwOnhIPYHQCod74ebM0Mv76nN_5X-Ud_C9VDZ6n4ucDUruUg5F4aFDi0Rm8f5_KtlSqyZL49HYnTcrKbPzrlPwGe5yZmOUVklajHRgNiy6-oljr2GEt0aCgYKAdISARESFQHGX2MiRe3-snCpajifm-rXAxbMjw0175';
const refreshToken =  '1//05qLZsqsCg8cLCgYIARAAGAUSNwF-L9IrXMberoaTBnBJ8OaOZbx3a5PiBg5bmAuOKrk3W7CkUu83jCRtViF9NeLXSPlZFTKwq3Q';
const scope = 'https://www.googleapis.com/auth/spreadsheets';
 

   // ฟังก์ชันสำหรับรีเฟรช access token เมื่อมีการคลิก
function refreshAccessToken() {


 if (!refreshToken) {
     console.error("No refresh token found.");
     return;
 }

 const url = "https://oauth2.googleapis.com/token";

 const params = new URLSearchParams();
 params.append("grant_type", "refresh_token");
 params.append("client_id", clientId);  // ใส่ clientId ของคุณที่นี่
 params.append("client_secret", clientSecret);  // ใส่ clientSecret ของคุณที่นี่
 params.append("refresh_token", refreshToken);

showLoading();																					 
 fetch(url, {
     method: 'POST',
     headers: {
         'Content-Type': 'application/x-www-form-urlencoded',
     },
     body: params
 })
 .then(response => {
     if (!response.ok) {
         throw new Error('Failed to refresh access token');
     }
     return response.json();
 })
 .then(data => {
     // เก็บ access token ใหม่
     sessionStorage.setItem("access_token", data.access_token);
     console.log("Access token refreshed:", data.access_token);
 })
 .catch(error => {
     console.error('Error refreshing access token:', error);
 })
 .finally(() => {
        hideLoading();
    });   
}

// ฟังก์ชันเพื่อตรวจสอบและรีเฟรช token เมื่อคลิก
function checkAndRefreshToken() {
 
 if (!access_token) {
     console.error("No access token found.");
     return;
 }

 console.log("Checking if access token needs refresh...");
 refreshAccessToken();  // เรียกฟังก์ชันรีเฟรชเมื่อคลิก
}


function searchData(){
const searchKeyElement = document.getElementById('searchKey');

if (searchKeyElement && searchKeyElement.value === '') {
  searchDataFromId();
} else {
searchDataKey();
}
}

function searchDataFromId() {
 const searchId = document.getElementById('idcard').textContent.trim(); // ดึงค่าจาก input และลบช่องว่าง
 
 const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}?key=${apiKey}`;

 fetch(url)
     .then(response => {
         if (!response.ok) {
             throw new Error('Network response was not ok');
         }
         return response.json();
     })
     .then(data => {
         const numb = document.getElementById("numb");
         const regisid = document.getElementById("registernumber");
         const name = document.getElementById("name");
       
         const card = document.getElementById("card");
         const depart = document.getElementById("depart");
         const age = document.getElementById("age");
         const birthday = document.getElementById("birthday");
         const program = document.getElementById("program");

         // ล้างผลลัพธ์ก่อนหน้า
         regisid.textContent = "";
         name.textContent = "";
        
         age.textContent = "";
         birthday.textContent = "";
         program.textContent = "";
         card.textContent = "";
         depart.textContent = "";

         let found = false; // ประกาศตัวแปร found เพื่อตรวจสอบว่าพบข้อมูลหรือไม่

         // ค้นหาและเก็บข้อมูลในตัวแปร searchResult
         if (data.values) {
             data.values.forEach(row => {
                 if (row[2] === searchId) {
                     // แสดงผลลัพธ์
                     regisid.textContent = row[0];
                     name.textContent = row[1];
                     
                     card.textContent = row[3];
                     depart.textContent = row[4];;
                     age.textContent = row[5];
                     birthday.textContent = row[6];
                     program.textContent = row[7];
                     found = true; // เปลี่ยนค่า found เป็น true

                     // เก็บค่าผลลัพธ์ในตัวแปร searchResult
                     searchResult = {
                         "regisid": row[0],
                         "name": row[1],
                         
                         "age": row[3],
                         "birthday": row[4],
                         "program": row[5]
                     };

                  setTimeout(() => {   
                     searchProgram();
                      }, 5000);
                     searchPrint();
                 }
             });
         }

         // ถ้าไม่พบข้อมูลที่ค้นหา
         if (!found) {
             alert("ไม่พบข้อมูลในระบบ");
         }
     })
     .catch(error => {
         console.error('Error:', error);
         alert("เกิดข้อผิดพลาดในการค้นหาข้อมูล");
     });
}   




function searchDataKey() {
    displayNextNumber();
 const searchKey = document.getElementById('searchKey').value.trim(); // ดึงค่าจาก input และลบช่องว่าง
 
 const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}?key=${apiKey}`;

 fetch(url)
     .then(response => {
         if (!response.ok) {
             throw new Error('Network response was not ok');
         }
         return response.json();
     })
     .then(data => {
         const numb = document.getElementById("numb");
         const regisid = document.getElementById("registernumber");
         const name = document.getElementById("name");
         const idcard = document.getElementById("idcard");
         const card = document.getElementById("card");
         const depart = document.getElementById("depart");
         const age = document.getElementById("age");
         const birthday = document.getElementById("birthday");
         const program = document.getElementById("program");

         // ล้างผลลัพธ์ก่อนหน้า
         regisid.textContent = "";
         name.textContent = "";
         idcard.textContent = "";
         age.textContent = "";
         birthday.textContent = "";
         program.textContent = "";
         card.textContent = "";
         depart.textContent = "";

         let found = false; // ประกาศตัวแปร found เพื่อตรวจสอบว่าพบข้อมูลหรือไม่

         // ค้นหาและเก็บข้อมูลในตัวแปร searchResult
         if (data.values) {
             data.values.forEach(row => {
                 if (row[2] === searchKey) {
                     // แสดงผลลัพธ์
                     regisid.textContent = row[0];
                     name.textContent = row[1];
                     idcard.textContent = row[2];
                     card.textContent = row[3];
                     depart.textContent = row[4];;
                     age.textContent = row[5];
                     birthday.textContent = row[6];
                     program.textContent = row[7];
                     found = true; // เปลี่ยนค่า found เป็น true

                     // เก็บค่าผลลัพธ์ในตัวแปร searchResult
                     searchResult = {
                         "regisid": row[0],
                         "name": row[1],
                         "idcard": row[2],
                         "age": row[3],
                         "birthday": row[4],
                         "program": row[5]
                     };

                  setTimeout(() => {   
                     searchProgram();
                      }, 5000);
                     searchPrint();
                 }
             });
         }

         // ถ้าไม่พบข้อมูลที่ค้นหา
         if (!found) {
             alert("ไม่พบข้อมูลในระบบ");
         }
     })
     .catch(error => {
         console.error('Error:', error);
         alert("เกิดข้อผิดพลาดในการค้นหาข้อมูล");
     });
}

function searchProgram() {
 
 const url1 = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet2}?key=${apiKey}`;
  showLoading();	 																						
 fetch(url1)
     .then(response => {
         if (!response.ok) {
             throw new Error('Network response was not ok');
         }
         return response.json();
     })
     .then(data => {
         const programdetail = document.getElementById('programdetail');
         const programName =  document.getElementById('program').textContent;
         programdetail.textContent = ""; // ล้างข้อมูลก่อนหน้า

         let found = false; // กำหนดค่า found เป็น false เริ่มต้น

         // ค้นหาโปรแกรมที่ตรงกับชื่อที่ให้มา
         if (data.values && data.values.length > 0) { // ตรวจสอบว่ามีข้อมูลใน data.values
             data.values.forEach(row => {
                 if (row[0] === programName) {
                     programdetail.innerHTML += `<p>- ${row[1]}</p>`; // แสดงรายละเอียดโปรแกรม
                     found = true; // เปลี่ยนค่า found เป็น true ถ้าพบโปรแกรม
                 }
             });
             hideLoading();
         }

         // หากไม่พบโปรแกรมที่ค้นหา
         if (!found) {
             programdetail.innerHTML = `<p>ไม่พบโปรแกรมที่ต้องการ</p>`;
         }
     })
     .catch(error => {
         console.error('Error:', error);
         const programdetail = document.getElementById('programdetail');
         programdetail.innerHTML = `<p>เกิดข้อผิดพลาดในการดึงข้อมูล</p>`;
     });
}


function loadAllRecords() {

 
 const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet10}?key=${apiKey}`;

 fetch(url)
     .then(response => {
         if (!response.ok) {
             throw new Error('Network response was not ok');
         }
         return response.json();
     })
     .then(data => {
         const resultDiv1 = document.getElementById('registeresult');

         resultDiv1.innerHTML = ''; // เคลียร์ผลลัพธ์ก่อนแสดงใหม่

         if (data.values && data.values.length > 0) { // ตรวจสอบว่ามีข้อมูลใน data.values
             data.values.forEach(row => {
                 resultDiv1.innerHTML += `
<tr>
 <th scope="col" class="text-center">${row[0]}</th>
 <td scope="col" colspan="2" class="text-center" style="font-family: sarabun;">${row[1]}</td>
 <td scope="col" colspan="4" style="font-family: sarabun;">${row[2]}</td>
 <td scope="col" class="text-center" style="font-family: sarabun;">${row[8]}</td>
 <td scope="col" class="text-center" style="font-family: sarabun;">${row[9]}</td>
</tr>
`;
              loadAllData();
             });
         } else {
             resultDiv1.innerHTML = `<tr><td colspan="8" class="text-center">ไม่พบข้อมูล</td></tr>`;
         }
     })
     .catch(error => {
         console.error('Error loading all records:', error);
         alert("เกิดข้อผิดพลาดในการดึงข้อมูล");
     });
}

function formatDateTime() {
 const now = new Date(); // ดึงวันที่และเวลาปัจจุบัน

 // ดึงค่าวัน, เดือน, ปี
 const day = String(now.getDate()).padStart(2, '0'); // เติมเลข 0 ถ้าวันน้อยกว่า 10
 const month = String(now.getMonth() + 1).padStart(2, '0'); // เดือนใน JavaScript เริ่มที่ 0
 const year = String(now.getFullYear()).slice(2); // เอาสองหลักสุดท้ายของปี

 // ดึงค่า ชั่วโมงและนาที
 const hours = String(now.getHours()).padStart(2, '0'); // เติมเลข 0 ถ้าชั่วโมงน้อยกว่า 10
 const minutes = String(now.getMinutes()).padStart(2, '0'); // เติมเลข 0 ถ้านาทีน้อยกว่า 10

 // รวมเป็นรูปแบบที่ต้องการ: dd/mm/yy, hh:mm
 return `${day}/${month}/${year}, ${hours}:${minutes}`;
}

// แสดงวันที่และเวลาปัจจุบัน
console.log(formatDateTime());

function updateDateTime() {
 const now = new Date();
 const date = now.toLocaleDateString('th-TH');
 const time = now.toLocaleTimeString('th-TH', { hour12: false }); // ใช้ 24 ชั่วโมง
 document.getElementById('datetime').textContent = `${date} ${time}`;
}

// เรียกใช้ updateDateTime ทุก 1 วินาที
setInterval(updateDateTime, 1000);


window.onload = function(){
loadAllRecords();
getNextNumber();
getNextSpecimenNumber();
updateDateTime();
loadAllData();
}
 







async function addRegistrationData() {
  await displayNextNumber();

  var number = document.getElementById('numb').textContent.trim();
  var regisid = document.getElementById('registernumber').textContent.trim();
  var name = document.getElementById('name').textContent.trim();
  var idcard = document.getElementById('idcard').textContent.trim();
  var sexage = document.getElementById('age').textContent.trim();
  var card = document.getElementById('card').textContent.trim();
  var depart = document.getElementById('depart').textContent.trim();
  var birth = document.getElementById('birthday').textContent.trim();
  var prog = document.getElementById('program').textContent.trim();
  var date = document.getElementById('datetime').textContent.trim();

  // สร้าง object ที่จะส่งไปยัง Google Sheets
  var data = {
    values: [[number, regisid, name, idcard, card, depart, sexage, birth, prog, date]]
  };

  checkAndRefreshToken(); // ตรวจสอบและรีเฟรช token

  // รอให้ access_token ถูกอัปเดตใน sessionStorage ก่อนทำการเพิ่มข้อมูล
  const accessToken = sessionStorage.getItem("access_token");

  // ตรวจสอบข้อมูล idcard ก่อน
  var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet3}?key=${apiKey}`; // แทนที่ด้วย API Key ของคุณ

  // ดึงข้อมูลทั้งหมดจากแผ่นงาน Google Sheets
  fetch(url, {
    method: 'GET',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': 'Bearer ' + accessToken
    }
  })
  .then(response => {
    if (!response.ok) {
      throw new Error('Network response was not ok: ' + response.statusText);
    }
    return response.json();
  })
  .then(sheetData => {
    // ตรวจสอบว่ามี idcard ซ้ำหรือไม่
    const rows = sheetData.values || [];
    const isDuplicate = rows.some(row => row[3] === idcard); // สมมุติว่า idcard อยู่ในคอลัมน์ที่ 4 (index 3)

    if (isDuplicate) {
      // แจ้งเตือนว่ามี idcard นี้ลงทะเบียนแล้ว
      Swal.fire({
        icon: 'error',
        title: 'Oops...',
        text: 'ID นี้ลงทะเบียนแล้ว!'
      });
    } else {
      // ถ้าไม่มี idcard ซ้ำ ให้ทำการโพสต์ข้อมูลใหม่
      var postUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet3}:append?valueInputOption=USER_ENTERED&key=${apiKey}`;

      fetch(postUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + accessToken
        },
        body: JSON.stringify(data)
      })
      .then(response => {
        if (!response.ok) {
          throw new Error('Network response was not ok: ' + response.statusText);
        }
        return response.json();
      })
      .then(data => {
        loadAllRecords(); // โหลดข้อมูลใหม่
        Swal.fire({
          position: "center",
          icon: "success",
          title: "ลงทะเบียนสำเร็จ",
          showConfirmButton: false,
          timer: 1500
        });
        addRegistrationDataInner(); // โหลดข้อมูลใหม่แม้เกิดข้อผิดพลาด
      })
      .catch(error => {
        console.error('Error:', error);
        Swal.fire({
          icon: 'error',
          title: 'Oops...',
          text: 'เกิดข้อผิดพลาดในการลงทะเบียน!'
        });
      });
    }
  })
  .catch(error => {
    console.error('Error fetching sheet data:', error);
    Swal.fire({
      icon: 'error',
      title: 'Oops...',
      text: 'เกิดข้อผิดพลาดในการตรวจสอบข้อมูล!'
    });
  });
}


async function addRegistrationDataInner() {

await  displayNextSpecimenNumber();
 var numr =  document.getElementById('specimenque').textContent.trim();
 var regisid = document.getElementById('registernumber').textContent.trim();
 var name = document.getElementById('name').textContent.trim();
 const type = "1ลงทะเบียน";
 const spec = "10"
 // สร้าง object ที่จะส่งไปยัง Google Sheets
 var data = {
     values: [[numr, regisid, name, spec, type]]
 };
 
 checkAndRefreshToken(); // ตรวจสอบและรีเฟรช token
 
 // รอให้ access_token ถูกอัปเดตใน sessionStorage ก่อนทำการเพิ่มข้อมูล
 const accessToken = sessionStorage.getItem("access_token");
 var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}:append?valueInputOption=USER_ENTERED&key=${apiKey}`; // แทนที่ด้วย API Key ของคุณ

 // ส่งข้อมูลไปยัง Google Sheets
 fetch(url, {
     method: 'POST',
    headers: {
 'Content-Type': 'application/json',
 'Authorization': 'Bearer ' + accessToken, // เพิ่มช่องว่างระหว่าง 'Bearer' และ accessToken
},

     body: JSON.stringify(data)
 })
 .then(response => {
     if (!response.ok) {
         throw new Error('Network response was not ok: ' + response.statusText);
     }
     return response.json();
 })

 .catch(error => {
     console.error('Error:', error);
     Swal.fire({
         icon: 'error',
         title: 'Oops...',
         text: 'เกิดข้อผิดพลาดในการลงทะเบียน!'
     });
     loadAllRecords(); // โหลดข้อมูลใหม่แม้เกิดข้อผิดพลาด
    
 });
}

 




function updateDate() {
const now = new Date();
const date = now.toLocaleDateString('th-TH');
const time = now.toLocaleTimeString('th-TH');
const current = `${date}`;
return current
       }

function searchPrint() {
 const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet5}?key=${apiKey}`;

 fetch(url)
     .then(response => {
         if (!response.ok) {
             throw new Error('Network response was not ok');
         }
         return response.json();
     })
     .then(data => {
         const sticker = document.getElementById('sticker');
         const number = document.getElementById('numb').textContent.trim(); // ใช้ trim() เพื่อหลีกเลี่ยงช่องว่าง
         const searchKey = document.getElementById('registernumber').textContent.trim();
         sticker.innerHTML = ''; // เคลียร์ก่อนแสดงใหม่

         let found = false; // ตัวแปรเพื่อตรวจสอบว่าพบข้อมูลหรือไม่

         // ค้นหาข้อมูลที่ตรงกับ searchKey
         data.values.forEach((row, index) => {
             if (row[0] === searchKey) {
                 found = true; // เปลี่ยนค่า found เป็น true

                 // สร้าง ID ที่ไม่ซ้ำกันสำหรับแต่ละสติ๊กเกอร์
                 const uniqueId = `barcode-${index}`;

                 sticker.innerHTML += `
                 <div class="sticker mb-2">
                     <div class="sticker-id mx-2 my-2"><span>${row[0]}</span></div>
                     <div class="sticker-right">
                         <div class="barcode-container mt-1">
                         <div class="mt-2 ms-2"><span class="libre-barcode-39-extended-regular">${row[1]}</span></div>
                         </div>
                         <div class="sticker-inner">
                             <div class="sticker-id"><span>${number}</span></div>
                             <div class="sticker-right">
                                 <div class="sticker-name ms-2">${row[3]}</div>
                                 <div class="sticker-name">
                                     <div class="sticker-method ms-2">${row[5]}</div>
                                     <div class="sticker-date mt-1">${updateDate()}</div>
                                 </div>
                                 <div class="sticker-name" style="font-size: 0.5rem">${row[4]}</div>
                             </div>
                         </div>
                     </div>
                 </div>
                 `;

                 const barcodeValue = row[1];
                 console.log(barcodeValue);
             }
         });

         // หากไม่พบข้อมูล ให้แจ้งเตือนหรือจัดการในกรณีไม่พบข้อมูล
         if (!found) {
             alert("ไม่พบข้อมูลที่ค้นหา");
         }
     })
     .catch(error => {
         console.error('Error fetching data:', error);
     });
}


function printSticker() {
 closeAlert();
 $('#sticker').css('display', 'block');
}

function closeModal() {
 $('#sticker').css('display', 'none');
}

function printSpecimen() {
 printResult();
 closeModal();
 clearPage();
    
}

function printResult() {
 const modalContent = document.querySelector('#sticker').innerHTML; 
 const originalContent = document.body.innerHTML;

 // แสดงเฉพาะเนื้อหาที่ต้องการพิมพ์
 document.body.innerHTML = modalContent;
 window.print();
 document.body.innerHTML = originalContent;
}





function clearPage() {
   var searchKey = document.getElementById('searchKey');
   var regisid = document.getElementById('registernumber');
   var name = document.getElementById('name');
   var idcard = document.getElementById('idcard');
   var age = document.getElementById('age');
   var birthday = document.getElementById('birthday');
   var program = document.getElementById('program');
   var programDetail = document.getElementById('programdetail');


   searchKey.innerHTML = '';
   regisid.textContent = '';
   name.textContent = '';
   idcard.textContent = '';
   age.textContent = '';
   birthday.textContent = '';
   program.textContent = '';
   programDetail.textContent = '';

}

function checkInputLength(){
    var input = document.getElementById('inputbar').value;
if (input.length === 8) {
     runFunction();
 }
}

async function runFunction() {
  try {
    await sendBarcode();                
    await displayNextSpecimenNumber();  
    await addRegistData();             
    await loadAllData();               
    await clearSpecimen();             
  } catch (error) {
    console.error("Error occurred in sequence:", error);
  }
}



function loadAllData() {
  return new Promise((resolve, reject) => {
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}?key=${apiKey}`;

 fetch(url)
     .then(response => {
         if (!response.ok) {
             throw new Error("Network response was not ok " + response.statusText);
         }
         return response.json();
     })
     .then(data => {
         const resultDiv1 = document.getElementById('specimenresult');
         resultDiv1.innerHTML = ''; // เคลียร์ผลลัพธ์ก่อนแสดงใหม่

         // ตรวจสอบว่ามีข้อมูลหรือไม่
         if (!data.values || data.values.length === 0) {
             resultDiv1.innerHTML = "<tr><td colspan='8' class='text-center'>ไม่พบข้อมูล</td></tr>";
             return; // ออกจากฟังก์ชันถ้าไม่มีข้อมูล
         }

         // เรียงลำดับข้อมูลจากคอลัมน์ [0] จากมากไปน้อย
         const sortedData = data.values.sort((a, b) => {
             const valueA = parseInt(a[0], 10); // แปลงค่าเป็นตัวเลข
             const valueB = parseInt(b[0], 10);
             return valueB - valueA; // เรียงจากมากไปน้อย
         });

         // แสดงข้อมูลใน resultDiv1
         sortedData.forEach(row => {
             resultDiv1.innerHTML += 
               `<tr>
                 <th scope="row" class="text-center">${row[0]}</th>
                 <td scope="col" colspan="2" class="text-center" style="font-family: sarabun;">${row[1] || 'N/A'}</td>
                 <td scope="col" colspan="6" class="text-align-start" style="font-family: sarabun;">${row[2] || 'N/A'}</td>
                 <td scope="col" class="text-center" style="font-family: sarabun;">${row[4] || 'N/A'}</td>
               </tr>`;
         });

         // เรียกใช้ loadAllCount() หลังจากแสดงผลข้อมูล
         loadAllCount();
     })
     .catch(error => {
         console.error("Error fetching data:", error);
         alert('เกิดข้อผิดพลาดในการโหลดข้อมูล');
     });
     
        resolve(); // เรียก resolve เมื่อประมวลผลเสร็จ
  });
}

function addRegistData() {
  return new Promise((resolve, reject) => {
    checkAndRefreshToken();
    var number1 = document.getElementById('specimenque').textContent.trim();
    var barinput = document.getElementById('inputbar').value.trim();
    var barcodenewid = document.getElementById('barregisterid').textContent.trim();
    var barcodename = document.getElementById('barname').textContent.trim();
    var inputmethodbar = barinput.slice(-2); // ดึง 2 ตัวท้ายของบาร์โค้ด 
    var barinputmethod = String(inputmethodbar);
    var specimen;

    switch (barinputmethod) {
        case "11":
            specimen = "PE";
            break;
        case "12":
            specimen = "EDTA";
            break;
        case "13":
            specimen = "Urine";
            break;
        case "14":
            specimen = "X Ray";
            break;
        case "15":
            specimen = "EKG";
            break;
        case "20":
            specimen = "naf";
            break;
        case "21":
            specimen = "Clot";
            break;
        case "16":
            specimen = "Audiogram";
            break;
        case "17":
            specimen = "Lung";
            break;
        case "18":
            specimen = "Eyes";
            break;
        case "19":
            specimen = "Muscle";
            break;
        default:
            specimen = "ไม่พบข้อมูล";
            break;
    }

    if (!barcodenewid || !barcodename || !barinputmethod || specimen === "ไม่พบข้อมูล") {
        alert("ข้อมูลไม่ครบถ้วน หรือไม่ถูกต้อง กรุณาตรวจสอบอีกครั้ง");
        return;
    }

    var data = {
        values: [[number1, barcodenewid, barcodename, barinputmethod, specimen]]
    };
    console.log(data);
    const accessToken = sessionStorage.getItem("access_token");
    var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}:append?valueInputOption=USER_ENTERED&key=${apiKey}`;

    fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + accessToken, // แก้ไขช่องว่างระหว่าง Bearer กับ token
        },
        body: JSON.stringify(data)
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Network response was not ok: ' + response.statusText);
        }
        return response.json();
    })
    .then(data => {
        console.log("Success:", data);
       

        
    })
    .catch(error => {
        console.error('Error:', error);
        alert("เกิดข้อผิดพลาดในการเพิ่มข้อมูล!");
    });

 resolve(); // เรียก resolve เมื่อประมวลผลเสร็จ
  });
}

function sendBarcode() {
  return new Promise((resolve, reject) => {
    var barcode = document.getElementById('inputbar').value.trim();
    var barcodeid = barcode.substring(0, 6); // เอา 6 ตัวแรกของบาร์โค้ดมา

    var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet3}?key=${apiKey}`;

    fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok: ' + response.statusText);
            }
            return response.json();
        })
        .then(data => {
            var records = data.values || [];
            var foundRecord = records.find(record => record[1] === barcodeid);

            var baridElement = document.getElementById('barregisterid');
            var barnameElement = document.getElementById('barname');

            baridElement.textContent = '';
            barnameElement.textContent = '';

            if (foundRecord) {
                baridElement.textContent = foundRecord[1];
                barnameElement.textContent = foundRecord[2];

               
            } else {
                alert('ไม่พบ ID นี้ในระบบ');
            }
        })
        .catch(error => {
            console.error('Error in fetching barcode data:', error);
            alert('เกิดข้อผิดพลาดในการค้นหาข้อมูล');
        });

 resolve(); // เรียก resolve เมื่อประมวลผลเสร็จ
  });
}


function loadAllCount() {
  return new Promise((resolve, reject) => {
 const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}?key=${apiKey}`;
checkAndRefreshToken();

 fetch(url)
     .then(response => {
         if (!response.ok) {
             throw new Error("Network response was not ok");
         }
         return response.json();
     })
     .then(data => {
         // ตรวจสอบว่ามีข้อมูลใน data.values หรือไม่
         if (!data.values || data.values.length === 0) {
             console.error('No data found.');
             return;
         }

         // ประกาศตัวนับนอกลูป
         let a = 0, b = 0, c = 0, d = 0, e = 0, f = 0, g = 0, h = 0, i = 0, j = 0, k = 0;

         data.values.forEach(row => {
             const barcodemethod = row[3]; // ตรวจสอบข้อมูลใน index 3

             // นับจำนวนการเกิดของแต่ละ `barcodemethod`
             switch (barcodemethod) {
                 case "11":
                     a++; 
                     break;
                 case "12":
                     b++; 
                     break;
                 case "13":
                     c++; 
                     break;
                 case "14":
                     d++; 
                     break;
                 case "15":
                     e++; 
                     break;
                 case "16":
                     f++; 
                     break;
                 case "17":
                     g++; 
                     break;
                 case "18":
                     h++; 
                     break;
                 case "19":
                     i++; 
                     break;
                 case "20":
                     j++; 
                     break;
                 case "21":
                     k++; 
                     break;
                 default:
                     console.log("Unrecognized barcode method:", barcodemethod);
                     break;
             }
         });

         // อัปเดตค่าใน HTML หลังจากประมวลผลเสร็จสิ้น
         document.getElementById('PE').textContent = a;      // พบแพทย์
         document.getElementById('EDTA').textContent = b;    // เจาะเลือด
         document.getElementById('urine').textContent = c;   // ปัสสาวะ
         document.getElementById('xray').textContent = d;    // X Ray
         document.getElementById('ekg').textContent = e;     // EKG
         document.getElementById('ear').textContent = f;     // Audiogram
         document.getElementById('lung').textContent = g;    // เป่าปอด
         document.getElementById('eye').textContent = h;     // ตา(ชีวอนามัย)
         document.getElementById('muscle').textContent = i;  // กล้ามเนื้อ
         document.getElementById('naf').textContent = j;     // NAF
         document.getElementById('clot').textContent = k;    // Clot
     })
     .catch(error => {
         console.error('Error fetching data:', error);
     });

resolve(); // เรียก resolve เมื่อประมวลผลเสร็จ
  });
}

function clearSpecimen() {
  return new Promise((resolve, reject) => {
 var barcode = document.getElementById('inputbar');
 barcode.value = '';
  
      resolve(); // เรียก resolve เมื่อประมวลผลเสร็จ
  });
}

function displayNextSpecimenNumber() {
 return new Promise((resolve, reject) => {
 const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet7}?key=${apiKey}`;
 checkAndRefreshToken(); // ตรวจสอบและรีเฟรช token ก่อนทำการ fetch

 return fetch(url)
     .then(response => {
         if (!response.ok) {
             throw new Error("Network response was not ok " + response.statusText);
         }
         return response.json();
     })
     .then(data => {
         const values = data.values;
         if (values && values.length > 0) {
             const lastNumber = values[values.length - 1][0];
             const autoNumber = document.getElementById('specimenque');
             autoNumber.innerText = parseInt(lastNumber) + 1; 
         } else {
             return 1; // ถ้าไม่มีข้อมูลในคอลัมน์ A ให้เริ่มที่ 1
         }
     })

 resolve(); // เรียก resolve เมื่อประมวลผลเสร็จ
  });
}

function getNextSpecimenNumber() {

 const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet7}?key=${apiKey}`;
 checkAndRefreshToken(); // ตรวจสอบและรีเฟรช token ก่อนทำการ fetch

 return fetch(url)
     .then(response => {
         if (!response.ok) {
             throw new Error("Network response was not ok " + response.statusText);
         }
         return response.json();
     })
     .then(data => {
         const values = data.values;
         if (values && values.length > 0) {
             const lastNumber = values[values.length - 1][0];
             const autoNumber = document.getElementById('specimenque');
             autoNumber.innerText = parseInt(lastNumber) + 1; 
         } else {
             return 1; // ถ้าไม่มีข้อมูลในคอลัมน์ A ให้เริ่มที่ 1
         }
     })

}

function addNewData(access_token) {
 // ดึงข้อมูลจาก input elements
 var newid = document.getElementById('newid').value.trim();
 var newname = document.getElementById('newname').value.trim();
 var newidcard = document.getElementById('newidcard').value.trim();
 var birthdate = document.getElementById('birthdate').value.trim();
 var newcard = document.getElementById('newcard').value.trim();
 var newdepart = document.getElementById('newdepart').value.trim();
 var newage = document.getElementById('newage').textContent.trim(); // ใช้ innerText หรือ textContent ให้เหมาะสมกับ HTML
 var newprogram = document.getElementById('newprogram').value.trim();

 // ตรวจสอบว่าข้อมูลทั้งหมดถูกกรอก
 if (!newid || !newname || !newidcard || !birthdate || !newcard || !newdepart || !newage || !newprogram) {
     alert('กรุณากรอกข้อมูลให้ครบถ้วน');
     return; // ออกจากฟังก์ชันถ้ามีข้อมูลไม่ครบ
 }

 var newRow = [newid, newname, newidcard, newcard, newdepart, newage, birthdate, newprogram];
 checkAndRefreshToken(); // ตรวจสอบและรีเฟรช OAuth token
 const accessToken = sessionStorage.getItem("access_token");
 const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}:append?valueInputOption=USER_ENTERED&key=${apiKey}`;

 const body = {
     "values": [newRow]
 };

 // ส่งข้อมูลไปที่ Google Sheets API ด้วย OAuth Token
 fetch(url, {
     method: "POST",
   headers: {
 'Content-Type': 'application/json',
 'Authorization': 'Bearer ' + accessToken, // เพิ่มช่องว่างระหว่าง 'Bearer' และ accessToken
},


     body: JSON.stringify(body)
 })
 .then(response => {
     if (!response.ok) {
         throw new Error("Network response was not ok " + response.statusText);
     }
     return response.json();
 })
 .then(data => {
     console.log("Data added successfully", data);
      buildSticker();
    
   
 })
 .catch(error => {
     console.error("Error adding data:", error);
     Swal.fire({
         icon: 'error',
         title: 'เกิดข้อผิดพลาด',
         text: 'ไม่สามารถเพิ่มข้อมูลได้!'
     });
 });

}


 



function openSheet() {
 $('#sheet').css('display', 'flex');
}

function closeSheet() {
 $('#sheet').css('display', 'none');
}

function getAlert() {
 $('#alert').css('display', 'flex');
}

function closeAlert() {
 $('#alert').css('display', 'none');
}

function closeNewRegister() {
 $('.modalregister').css('display', 'none');
}

function openNewRegister() {
clearPage();
clearRegisterPage();
 $('.modalregister').css('display', 'block');
}

function openSpecimen() {
 $('.modalspecimen').css('display', 'flex');
}

function closeSpecimen() {
 $('.modalspecimen').css('display', 'none');
}


 async function displayNextNumber() {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet4}?key=${apiKey}`;
 checkAndRefreshToken(); // ตรวจสอบและรีเฟรช token ก่อนทำการ fetch
 return fetch(url)
     .then(response => {
         if (!response.ok) {
             throw new Error("Network response was not ok " + response.statusText);
         }
         return response.json();
     })
     .then(data => {
         const values = data.values;
         if (values && values.length > 0) {
             const lastNumber = values[values.length - 1][0];
             const autoNumb =  document.getElementById('numb');
             autoNumb.innerText = parseInt(lastNumber) + 1 ;
         } else {
             const autoNumb =  document.getElementById('numb');
             autoNumb.innerText = 1 ;
         }
     })
     .catch(error => {
         console.error('Error fetching data:', error);
         throw error; // ส่งต่อ error ไปยัง function ที่เรียกใช้
     });
}
 

  function getNextNumber() {
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet4}?key=${apiKey}`;
  checkAndRefreshToken(); // ตรวจสอบและรีเฟรช token ก่อนทำการ fetch
  return fetch(url)
     .then(response => {
         if (!response.ok) {
             throw new Error("Network response was not ok " + response.statusText);
         }
         return response.json();
     })
     .then(data => {
         const values = data.values;
         if (values && values.length > 0) {
             const lastNumber = values[values.length - 1][0];
             const autoNumb =  document.getElementById('numb');
             autoNumb.innerText = parseInt(lastNumber) + 1 ;
         } else {
             const autoNumb =  document.getElementById('numb');
             autoNumb.innerText = 1 ;
         }
     })
     .catch(error => {
         console.error('Error fetching data:', error);
         throw error; // ส่งต่อ error ไปยัง function ที่เรียกใช้
     });
}
 
 

 
function calculateAge() {
var birthdate = document.getElementById('birthdate').value;
 
 if (birthdate) {
   // แปลงวันที่เป็น Date object
   var birthDateObj = new Date(birthdate);
   var today = new Date(); // วันที่ปัจจุบัน
   var age = today.getFullYear() - birthDateObj.getFullYear(); // คำนวณอายุคร่าว ๆ

   // ตรวจสอบเดือนและวัน เพื่อให้แน่ใจว่าคำนวณอายุถูกต้อง (เผื่อยังไม่ถึงวันเกิดปีนี้)
   var monthDiff = today.getMonth() - birthDateObj.getMonth();
   if (monthDiff < 0 || (monthDiff === 0 && today.getDate() < birthDateObj.getDate())) {
     age--; // ถ้าเดือนนี้ยังไม่ถึงวันเกิด ให้ลบอายุออก 1 ปี
   }
   
   // แสดงผลใน <span> ที่มี id="newage"
   document.getElementById('newage').textContent = age;
 } else {
   // ถ้าไม่มีวันเดือนปีเกิด ให้ล้างค่าใน <span>
   document.getElementById('newage').textContent = '';
 }
}


function loadRegister() {
 const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet4}?key=${apiKey}`;

 fetch(url)
     .then(response => {
         if (!response.ok) {
             throw new Error('Network response was not ok: ' + response.statusText);
         }
         return response.json();
     })
     .then(data => {
         const registerCount = (data.values && data.values.length > 0) ? data.values.length : 0; // ตรวจสอบจำนวนข้อมูล
         document.getElementById('register').textContent = registerCount; // นับจำนวนแถว
     })
     
     .catch(error => {
         console.error('Error fetching data:', error);
         alert('เกิดข้อผิดพลาดในการโหลดจำนวนทะเบียน');
     });
}





function getCurrentDateTime() {
 const now = new Date();
 return now.toLocaleString(); // Date and time format
}

function updateDateTime() {
 const now = new Date();
 const date = now.toLocaleDateString('th-TH');
 const time = now.toLocaleTimeString('th-TH');
 document.getElementById('datetime').innerHTML = `${date} ${time}`;
}

// เรียกฟังก์ชันทุกวินาทีเพื่ออัปเดตวันที่และเวลา
setInterval(updateDateTime, 1000);

function openSheet() {
 $(".modalsheet").css('display', 'block');
}
function closeRegist(){
 clearRegisterPage();
 $(".modalregister").css('display', 'none');
}    
function closeSheet() {
 $(".modalsheet").css('display', 'none');
}

function openSearch() {
 $(".modalsearch").css('display', 'block');
}

function closeSearch() {
 $(".modalsearch").css('display', 'none');
}

function buildSticker() {
    const program = document.getElementById('newprogram').value.trim();
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet8}?key=${apiKey}`;

    fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
            console.log('Data received from API:', data);

            if (data.values) {
                const matchingRows = data.values.filter(row => row[0] === program);

                console.log('Matching rows:', matchingRows);

                if (matchingRows.length === 0) {
                    alert('ไม่พบข้อมูลที่ตรงกับโปรแกรม');
                    return;
                }

                matchingRows.forEach(row => {
                    const n1 = document.getElementById('newidcard').value.trim();          
                    const n2 = document.getElementById('idcard');
                    n2.innerText = n1;

                    const method = row[2] || 'Unknown method';  // ตรวจสอบว่ามีค่าไหม
                    const methodid = row[3] || 'Unknown methodid'; // ตรวจสอบว่ามีค่าไหม
                    const custom = row[4] || 'Unknown custom'; // ตรวจสอบว่ามีค่าไหม
                    const regisidInput = document.getElementById("newid");
                    const nameInput = document.getElementById("newname");
                    const programInput = document.getElementById("newprogram");

                    if (!regisidInput || !nameInput || !programInput) {
                        console.error('ไม่พบ element newid, newname หรือ newprogram');
                        alert('ไม่พบข้อมูลการลงทะเบียน');
                        return;
                    }

                    const regisid = regisidInput.value;
                    const name = nameInput.value;

                    if (!regisid || !name) {
                        console.error('ไม่สามารถดึงค่า regisid หรือ name');
                        alert('กรุณากรอกข้อมูลให้ครบ');
                        return;
                    }

                    const barcodesticker = "*" + String(regisid) + String(methodid) + "*";
                    const stickerid = String(regisid) + String(program);

                    const dataToSave = {
                        values: [[regisid, barcodesticker, stickerid, name, custom, method]]
                    };

                    checkAndRefreshToken();
                    const accessToken = sessionStorage.getItem("access_token");

                    const appendUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet9}:append?valueInputOption=USER_ENTERED&key=${apiKey}`;

                    fetch(appendUrl, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'Authorization': 'Bearer ' + accessToken,
                        },
                        body: JSON.stringify(dataToSave)
                    })
                    .then(response => {
                        if (!response.ok) {
                            throw new Error('Network response was not ok');
                        }
                        return response.json();
                    })
                    .then(result => {
                        console.log('Data saved successfully:', result);
                    })
                    .catch(error => {
                        console.error('เกิดข้อผิดพลาดในการบันทึกข้อมูล:', error);
                        alert('เกิดข้อผิดพลาดในการบันทึกข้อมูล');
                    });
                });
            }
        })
        .catch(error => {
            console.error('เกิดข้อผิดพลาดในการดึงข้อมูล:', error);
        });

    Swal.fire({
        position: 'center',
        icon: 'success',
        title: 'เพิ่มข้อมูลสำเร็จ',
        showConfirmButton: false,
        timer: 1500
    });

    setTimeout(() => { 
        clearRegisterPage(); 
    }, 1000);
    closeNewRegister();
}



function printSticker() {
closeAlert();
$('#sticker').css('display', 'block')

}

function printSspecimen() {
printResult();
closeModal();
 clearPage();
   closeModal();
    displayNextNumber();
     displayNextSpecimenNumber();
    
 

}

function printResult() {
  var modalContent = document.querySelector('#sticker').innerHTML; 
   var originalContent = document.body.innerHTML;

   
     document.body.innerHTML = modalContent;
     window.print();
      document.body.innerHTML = originalContent;
    

   
}


function clearPage() {
     var searchKey = document.getElementById('searchKey');
     var regisid = document.getElementById('registernumber');
     var name = document.getElementById('name');
     var idcard = document.getElementById('idcard');
     var age = document.getElementById('age');
     var birthday = document.getElementById('birthday');
     var program = document.getElementById('program');
     var programDetail = document.getElementById('programdetail');
     var card = document.getElementById('card');
     var depart = document.getElementById('depart');
     searchKey.value = '';
     regisid.textContent = '';
     name.textContent = '';
     idcard.textContent = '';
     age.textContent = '';
     birthday.textContent = '';
     program.textContent = '';
     programDetail.textContent = '';
     card.textContent = '';
     depart.textContent = '';
}


function clearRegisterPage() {
   
     var newregisid = document.getElementById('newid');
     var newname = document.getElementById('newname');
     var newidcard = document.getElementById('newidcard');
     var newage = document.getElementById('newage');
     var newbirthday = document.getElementById('birthdate');
     var newprogram = document.getElementById('newprogram');
     var newcard = document.getElementById('newcard');
     var newdepart = document.getElementById('newdepart');

     newdepart.value = '';
     newregisid.value = '';
     newname.value = '';
     newidcard.value = '';
     newage.textContent = '';
     newbirthday.textContent = '';
     newprogram.value = '';
     newcard.value = '';
}
