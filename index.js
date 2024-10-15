<script src="api.js"></script>
window.onload = function(){
    loadAllRecords();
    displayNextNumber();
    updateDateTime();
    loadAllData();
}



function searchData() {
    const searchKey = document.getElementById('searchKey').value;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}?key=${apiKey}`;

    fetch(url)
        .then(response => response.json())
        .then(data => {

            const numb = document.getElementById("numb");
            const regisid = document.getElementById("registernumber");
            const name = document.getElementById("name");
            const idcard = document.getElementById("idcard");
            const age = document.getElementById("age");
            const birthday = document.getElementById("birthday");
            const program = document.getElementById("program");

            // Clear previous results

            // ค้นหาและเก็บข้อมูลในตัวแปร searchResult
            data.values.forEach(row => {
                if (row[2] === searchKey) {
                    // แสดงผลลัพธ์
                    regisid.innerHTML = row[0];
                    name.innerHTML = row[1];
                    idcard.innerHTML = row[2];
                    age.innerHTML = row[3];
                    birthday.innerHTML = row[4];
                    program.innerHTML = row[5];
                    found = true;

                    // เก็บค่าผลลัพธ์ในตัวแปร searchResult
                    searchResult = {
                        "regisid": row[0],
                        "name": row[1],
                        "idcard": row[2],
                        "age": row[3],
                        "birthday": row[4],
                        "program": row[5]
                    };

                    // เรียกใช้งานฟังก์ชันอื่นพร้อมส่งข้อมูล
                    searchProgram(row[5]);
                    searchPrint();
                }
            });

            if (!found) {
                program.innerHTML = '<p>ไม่พบ ID นี้</p>';
            }
        })
        .catch(error => {
            console.error('Error:', error);
        });
}




function searchProgram(programName) {
    const url1 = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet2}?key=${apiKey}`;

    fetch(url1)
        .then(response => response.json())
        .then(data => {
            const programdetail = document.getElementById('programdetail');
            programdetail.innerHTML = ""; // Clear previous data

            let found = false;

            // Search for matching program
            data.values.forEach(row => {
                if (row[0] === programName) {
                    programdetail.innerHTML += `<p>-${row[1]}</p>`; // Show program detail
                    found = true;
                }
            });

            if (!found) {
                programdetail.innerHTML = '<p>ไม่พบข้อมูลโปรแกรม</p>'; // Show message if no data found
            }
        })
        .catch(error => {
            console.error('Error:', error);
        });
}



function getNextNumber() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet4}?key=${apiKey}`;

    return fetch(url)
        .then(response => response.json())
        .then(data => {
            const values = data.values;
            if (values && values.length > 0) {
                const lastNumber = values[values.length - 1][0];
                return parseInt(lastNumber) + 1; // Next number
            } else {
                return 1; // ถ้าไม่มีข้อมูลในคอลัมน์ A
            }
        })
        .catch(error => {
            console.error('Error fetching data:', error);
        });
}





 
function displayNextNumber() {
    getNextNumber().then(nextNumber => {
        // แสดงค่าใน element ที่มี id เป็น numb
        document.getElementById('numb').innerHTML = nextNumber;
    }).catch(error => {
        console.error('Error displaying next number:', error);
    });
}

function loadAllRecords() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet3}?key=${apiKey}`;

    fetch(url)
        .then(response => response.json())
        .then(data => {
            const resultDiv1 = document.getElementById('registeresult');




            resultDiv1.innerHTML = ''; // เคลียร์ผลลัพธ์ก่อนแสดงใหม่

            data.values.forEach(row => {

                resultDiv1.innerHTML += ` 
<tr>
<th scope="col" class="text-center">${row[0]}</th>
                    <td scope="col" colspan="2" class="text-center" style="font-family: sarabun;">${row[1]}</th>
                    <td scope="col" colspan="4"  style="font-family: sarabun;">${row[2]}</th>
                    <td scope="col" class="text-center" style="font-family: sarabun;">${row[6]}</th>
                    <td scope="col" class="text-center" style="font-family: sarabun;">${row[7]}</th>
 
 
 
</tr>
`;
            }
            );
        })

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
    const time = now.toLocaleTimeString('th-TH');
    document.getElementById('datetime').innerHTML = `${date} ${time}`;
            }

     setInterval(updateDateTime, 1000);









     function addRegistrationData() {
      // ดึงข้อมูลจาก HTML elements
      var numb = document.getElementById('numb').textContent.trim();
      var regisid = document.getElementById('registernumber').textContent.trim();
      var name = document.getElementById('name').textContent.trim();
      var idcard = document.getElementById('idcard').textContent.trim();
      var sexage = document.getElementById('age').textContent.trim();
      var birth = document.getElementById('birthday').textContent.trim();
      var prog = document.getElementById('program').textContent.trim();
      var date = document.getElementById('datetime').textContent.trim();
    
      // สร้าง object ที่จะส่งไปยัง Google Sheets
      var data = {
        values: [[numb, regisid, name, idcard, sexage, birth, prog, date]]
      };
    
     
      var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}:append?valueInputOption=USER_ENTERED&key=${apiKey}`; // แทนที่ด้วย API Key ของคุณ
    
      // ส่งข้อมูลไปยัง Google Sheets
      fetch(url, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': 'Bearer ' + access_token
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
        loadAllData(); // เรียกใช้ฟังก์ชันเพื่อโหลดข้อมูลใหม่
        Swal.fire({
          position: "center",
          icon: "success",
          title: "ลงทะเบียนสำเร็จ",
          showConfirmButton: false,
          timer: 1500
        });
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
    




function updateDate() {
  const now = new Date();
  const date = now.toLocaleDateString('th-TH');
  const time = now.toLocaleTimeString('th-TH');
  const current = `${date}`;
  return current
          }

function searchPrint(){

  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet5}?key=${apiKey}`;

  
 
  fetch(url)
        .then(response => response.json())
        .then(data => {
      var sticker = document.getElementById('sticker');
      var number = document.getElementById('numb').textContent;
      var searchKey = document.getElementById('registernumber').textContent.trim();
      sticker.textContent = '';
      
      data.values.forEach(row => {
        if (row[0] === searchKey){

       
          sticker.innerHTML+= 
                              ` <div class="sticker mb-2">
      <div class="sticker-id mx-2 my-2"><span>${row[0]}</span></div>

      <div class="sticker-right">
         <div class="barcode-container mt-1"> <span><svg id="barcode"></svg></span></div>
          
         <div class="sticker-inner">
              <div class="sticker-id">
                <span>${number}</span>
              </div>
            <div class="sticker-right">
              <div class="sticker-name">${row[3]}</div>
              <div class="sticker-name"><div class="sticker-method">${row[5]}</div><div class="sticker-date">${updateDate()}</div></div>
              <div class="sticker-name"style="font-size: 0.5rem">${row[4]}</div>
         </div>
       </div>
          
  </div>
  `;
  var barcodevalue = row[1];
  console.log(barcodevalue);
  
JsBarcode("#barcode",barcodevalue, {
    
      format: "CODE39",
      margin: 0,
      padding: 0,
      width: 1,
      height: 40,
      displayValue: false
});
     
        }});

        });
      
        
      }
         
      
    

function printSticker() {
closeAlert();
$('#sticker').css('display', 'block')

}

function closeModal() {

$('#sticker').css('display', 'none')

}

function printSspecimen() {
 printResult();
 closeModal();
  clearPage();
    closeModal();
     
     
  

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


      searchKey.innerHTML = '';
      regisid.textContent = '';
      name.textContent = '';
      idcard.textContent = '';
      age.textContent = '';
      birthday.textContent = '';
      program.textContent = '';
      programDetail.textContent = '';

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





function loadAllData() {
    
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}?key=${apiKey}`;
  
    fetch(url)
        .then(response => response.json())
        .then(data => {
            const resultDiv1 = document.getElementById('specimenresult');




            resultDiv1.innerHTML = ''; // เคลียร์ผลลัพธ์ก่อนแสดงใหม่

            data.values.forEach(row =>{
        resultDiv1.innerHTML += 
          `<tr>
            <th scope="col" class="text-center">${row[0]}</th>
            <td scope="col" colspan="2" class="text-center" style="font-family: sarabun;">${row[1]}</td>
            <td scope="col" colspan="6" class="text-align-start" style="font-family: sarabun;">${row[2]}</td>
            <td scope="col" class="text-center" style="font-family: sarabun;">${row[4]}</td>
           
          </tr>`;
      });
    loadAllCount();
     })}
        
    
     function addNewData(access_token) {
      var newid = document.getElementById('newid').value.trim();
      var newname = document.getElementById('newname').value.trim();
      var newidcard = document.getElementById('newidcard').value.trim();
      var birthdate = document.getElementById('birthdate').value.trim();
      var newage = document.getElementById('newage').textContent.trim();
      var newprogram = document.getElementById('newprogram').value.trim();
  
      var newRow = [newid, newname, newidcard, birthdate, newage, newprogram];
  
      const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}:append?valueInputOption=USER_ENTERED&key=${apiKey}`;
  
      const body = {
          "values": [newRow]
      };
  
      // ส่งข้อมูลไปที่ Google Sheets API ด้วย OAuth Token
      fetch(url, {
          method: "POST",
          headers: {
              "Content-Type": "application/json",
              "Authorization": `Bearer ${access_token}`, // ใช้ OAuth token ใน header
          },
          body: JSON.stringify(body)
      })
      .then(response => response.json())
      .then(data => {
          console.log("Data added successfully", data);
          closeNewRegister();
      })
      .catch(error => {
          console.error("Error:", error);
      });
  }
  
    



 function openSheet(){
   $('#sheet').css('display','flex')
}

function closeSheet(){
  $('#sheet').css('display','none')
}
function getAleart(){
  $('#alert').css('display','flex')
}


function closeAlert(){
  $('#alert').css('display','none')
}


function closeNewRegister() {
   $('.modalregister').css('display', 'none')
}


function openNewRegister() {
   $('.modalregister').css('display', 'block')
}
function openSpecimen() {
   $('.modalspecimen').css('display', 'flex')
}
function closeSpecimen() {
   $('.modalspecimen').css('display', 'none')
}

 function checkInputLength() {
            const input = document.getElementById('inputbar').value;

            // ตรวจสอบความยาวของ input
            if (input.length === 10) {
               
                runFunction();
            }
        }

        function runFunction() {
          
            sendBarcode(event);
        }

        function sendBarcode(event) {
          event.preventDefault(); // ป้องกันการรีโหลดหน้าเว็บ
        
          var barcode = document.getElementById('inputbar').value;
          var barcodeid = barcode.substring(0, 8); // เอา 8 ตัวแรกของบาร์โค้ดมา
        
          // URL สำหรับ Google Sheets API
         
          var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}?key=YOUR_API_KEY`; // แทนที่ด้วย API Key ของคุณ
        
          // ส่งคำขอ GET เพื่อตรวจสอบข้อมูลใน Google Sheets
          fetch(url)
            .then(response => {
              if (!response.ok) {
                throw new Error('Network response was not ok: ' + response.statusText);
              }
              return response.json();
            })
            .then(data => {
              var records = data.values || []; // ดึงข้อมูลจาก response
              var foundRecord = null;
        
              // ค้นหาข้อมูลที่ตรงกับ barcodeid
              records.forEach(function(record) {
                if (record[0] === barcodeid) { // สมมติว่า barid อยู่ในคอลัมน์แรก
                  foundRecord = {
                    barid: record[0], // barid
                    barname: record[1] // barname
                  };
                }
              });
        
              var baridElement = document.getElementById('barregisterid');
              var barnameElement = document.getElementById('barname');
        
              baridElement.textContent = '';
              barnameElement.textContent = '';
        
              if (foundRecord) {
                baridElement.textContent = foundRecord.barid;
                barnameElement.textContent = foundRecord.barname;
        
                var id = baridElement.textContent;
                var name = barnameElement.textContent;
                console.log(id, name);
                addRegistData();
              } else {
                alert('ไม่พบ ID นี้ในระบบ');
              }
            })
            .catch(error => {
              console.error('Error:', error);
              alert('เกิดข้อผิดพลาดในการค้นหาข้อมูล');
            });
        }
        


function addRegistData() {
  var barinput = document.getElementById('inputbar').value.trim();
  var barcodenewid = document.getElementById('barregisterid').textContent.trim();
  var barcodename = document.getElementById('barname').textContent.trim();
  var barinputmethod = barinput.slice(-2); // ดึง 2 ตัวสุดท้าย
  var specimen; // ประกาศตัวแปร specimen
  
  // ตรวจสอบค่า barcodemethod แล้วกำหนดค่าให้ specimen
  switch (barinputmethod) {
    case '11':
      specimen = "พบแพทย์";
      break;
    case '12':
      specimen = "เจาะเลือด";
      break;
    case '13':
      specimen = "ปัสสาวะ";
      break;
    case '14':
      specimen = "X Ray";
      break;
    case '15':
      specimen = "EKG";
      break;
    case '16':
      specimen = "Audiogram";
      break;
    case '17':
      specimen = "เป่าปอด";
      break;
    case '18':
      specimen = "ตา(ชีวอนามัย)";
      break;
    case '19':
      specimen = "Muscle";
      break;
    default:
      specimen = "ไม่พบข้อมูล";
      break;
  }
 
  if (!barcodenewid || !barcodename || !barinputmethod || !specimen) {
    alert("ข้อมูลไม่ครบถ้วน กรุณาตรวจสอบอีกครั้ง");
    return;
  }

  // สร้าง object ที่จะส่งไปยัง Google Sheets
  var data = {
    values: [[barcodenewid, barcodename, barinputmethod, specimen]]
  };

  
  
  var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}:append?valueInputOption=USER_ENTERED&key=${apiKey}`; // แทนที่ด้วย API Key ของคุณ

  // ส่งข้อมูลไปยัง Google Sheets
  fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      "Authorization": `Bearer ${access_token}`,
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
    loaddAllSpecimen(); // เรียกฟังก์ชันเพื่อโหลดข้อมูลใหม่
  })
  .catch(error => {
    console.error('Error:', error);
    alert("เกิดข้อผิดพลาดในการเพิ่มข้อมูล!");
  });

  // ล้างค่าใน input bar
  clearSpecimen();
  addDataToSheet();
}

function clearSpecimen(){
    var barcode = document.getElementById('inputbar');
    barcode.value = '';
}

function addDataToSheet(data) {
 
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}:append?valueInputOption=USER_ENTERED&key${apiKey}`; // แทนที่ด้วย API Key ของคุณ

  fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(data)
  })
  .then(response => response.json())
  .then(result => {
    console.log('Data added:', result);
    // ดำเนินการอื่นๆ ถ้าจำเป็น
  })
  .catch(error => {
    console.error('Error adding data:', error);
  });
}

// ฟังก์ชันสำหรับอัปเดตข้อมูลใน Google Sheets
function updateSheet(sheetName, column, data) {
 
  const range = `${sheetName}!${String.fromCharCode(64 + column)}2`; // ใช้คอลัมน์ที่ระบุ
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}?valueInputOption=USER_ENTERED&key=${apiKey}`; // แทนที่ด้วย API Key ของคุณ

  fetch(url, {
    method: 'PUT',
    headers: {
      'Content-Type': 'application/json'
    },
    body: JSON.stringify(data)
  })
  .then(response => response.json())
  .then(result => {
    console.log('Data updated:', result);
    // ดำเนินการอื่นๆ ถ้าจำเป็น
  })
  .catch(error => {
    console.error('Error updating data:', error);
  });
}


function loadAllCount() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}?key=${apiKey}`;

    fetch(url)
        .then(response => response.json())
        .then(data => {
            // ประกาศตัวนับนอกลูป
            let a = 0, b = 0, c = 0, d = 0, e = 0, f = 0, g = 0, h = 0, i = 0, j = 0, k = 0;

            data.values.forEach(row => {
                const barcodemethod = row[3]; 
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
                        j++;  // แก้ไขตรงนี้ เดิมนับค่า h แทน j
                        break;
                    case "21":
                        k++;  // แก้ไขตรงนี้ เดิมนับค่า i แทน k
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

    loadRegister(); 
    clearSpecimen();  
}


function loadRegister() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet4}?key=${apiKey}`;

    fetch(url)
        .then(response => response.json())
        .then(data => {
            if (data.values && data.values.length > 0) {
                let z = 0;
                data.values.forEach(row => {
                    z++; // Count the number of rows
                });

                const regis = document.getElementById('register');
                regis.textContent = z; 
            } else {
                console.error('No data found');
            }
        })
        .catch(error => {
            console.error('Error fetching data:', error);
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



function openSheet(){
   $(".modalsheet").css('display','block')
}

function closeSheet(){
   $(".modalsheet").css('display','none')
}

function openSearch(){
   $(".modalsearch").css('display','block')
}

function closeSearch(){
   $(".modalsearch").css('display','none')
}


