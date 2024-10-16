const client_Id = '168121174551-ij5g6b5l20kjk89n69kk4h7i518vvqrb.apps.googleusercontent.com';
const client_secret = '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';
const apiKey = 'AIzaSyAacPAesETbnRTARngdDeEffxrqBxVmGcg';
const redirectUri = 'https://rbmedical.github.io/register/';
const scope = 'https://www.googleapis.com/auth/spreadsheets';
const spreadsheetId = '1_aUWV9uDvVn_WBs25ZsHtVLilUYB9iNP87yadjSbHsw';
const rangesheet1 = 'data!A2:ZZ'; 
const rangesheet2 = 'program!A2:ZZ';
const rangesheet3 = 'register!A2:ZZ';
const rangesheet4 = 'register!A:A';
const rangesheet5 = 'sticker!A2:ZZ';
const rangesheet6 = 'specimencount!A2:ZZ';
let access_token; 

function initialize() {
    console.log(client_Id);
    console.log(apiKey);
}

function useAccessToken(token) {
    access_token = token; // ตั้งค่า access_token ที่ได้รับ
    console.log(access_token); // แสดงค่า access_token
}

window.onload = function() {
    initialize();
    const urlParams = new URLSearchParams(window.location.search);
    const code = urlParams.get('code');

    if (code) {
        exchangeCodeForToken(code);
    } else {
        requestAccessToken();
    }
}

function requestAccessToken() {
    // ขั้นตอนที่ 1: ขอ Authorization Code
    const authUrl = `https://accounts.google.com/o/oauth2/v2/auth?` +
        `client_id=${client_Id}&` + // แก้ไขจาก clientId เป็น client_Id
        `redirect_uri=${redirectUri}&` +
        `response_type=code&` +
        `scope=${encodeURIComponent(scope)}`;

    // ส่งผู้ใช้ไปยังหน้าการรับรอง
    window.location.href = authUrl;
}

function exchangeCodeForToken(code) {
    // ขั้นตอนที่ 2: แลกเปลี่ยน Authorization Code เป็น Access Token
    const tokenUrl = 'https://oauth2.googleapis.com/token';
    
    const data = new URLSearchParams();
    data.append('code', code);
    data.append('client_id', client_Id); // แก้ไขจาก clientId เป็น client_Id
    data.append('client_secret', client_secret); // ใช้ client_secret ที่ประกาศไว้
    data.append('redirect_uri', redirectUri);
    data.append('grant_type', 'authorization_code');

    fetch(tokenUrl, {
        method: 'POST',
        body: data,
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    })
    .then(response => response.json())
    .then(data => {
        console.log(data); // แสดงผลลัพธ์ใน console
        if (data.access_token) {
            useAccessToken(data.access_token);
            loadData(); // เรียกใช้ฟังก์ชันเพื่อโหลดข้อมูลหลังจากได้ access token
        } else {
            console.error('Failed to get access token:', data);
        }
    })
    .catch(error => console.error('Error fetching access token:', error));
}

function loadData() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}?access_token=${access_token}`;
    
    fetch(url)
        .then(response => response.json())
        .then(data => console.log(data))
        .catch(error => console.error('Error accessing Google Sheets API:', error));
}

function searchData() {
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

            let found = false; // ประกาศตัวแปร found เพื่อตรวจสอบว่าพบข้อมูลหรือไม่

            // ค้นหาและเก็บข้อมูลในตัวแปร searchResult
            if (data.values) {
                data.values.forEach(row => {
                    if (row[2] === searchKey) {
                        // แสดงผลลัพธ์
                        regisid.textContent = row[0];
                        name.textContent = row[1];
                        idcard.textContent = row[2];
                        age.textContent = row[3];
                        birthday.textContent = row[4];
                        program.textContent = row[5];
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

                        // เรียกใช้งานฟังก์ชันอื่นพร้อมส่งข้อมูล
                        searchProgram(row[5]);
                        searchPrint();
                    }
                });
            }

            // ถ้าไม่พบข้อมูลที่ค้นหา
            if (!found) {
                alert("ไม่พบข้อมูลที่ต้องการ");
            }
        })
        .catch(error => {
            console.error('Error:', error);
            alert("เกิดข้อผิดพลาดในการค้นหาข้อมูล");
        });
}

function searchProgram(programName) {
    const url1 = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet2}?key=${apiKey}`;

    fetch(url1)
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
            const programdetail = document.getElementById('programdetail');
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

function getNextNumber() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet4}?key=${apiKey}`;

    return fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
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
        document.getElementById('numb').textContent = nextNumber;
    }).catch(error => {
        console.error('Error displaying next number:', error);
    });
}

function loadAllRecords() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet3}?key=${apiKey}`;

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
    <td scope="col" class="text-center" style="font-family: sarabun;">${row[6]}</td>
    <td scope="col" class="text-center" style="font-family: sarabun;">${row[7]}</td>
</tr>
`;
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
    
     
      var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet3}:append?valueInputOption=USER_ENTERED&key=${apiKey}`; // แทนที่ด้วย API Key ของคุณ
    
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
          loadAllData();
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
            const number = document.getElementById('numb').textContent;
            const searchKey = document.getElementById('registernumber').textContent.trim();
            sticker.innerHTML = ''; // เคลียร์ก่อนแสดงใหม่

            let found = false; // ตัวแปรเพื่อตรวจสอบว่าพบข้อมูลหรือไม่

            data.values.forEach(row => {
                if (row[0] === searchKey) {
                    found = true; // เปลี่ยนค่า found เป็น true

                    sticker.innerHTML += `
                    <div class="sticker mb-2">
                        <div class="sticker-id mx-2 my-2"><span>${row[0]}</span></div>
                        <div class="sticker-right">
                            <div class="barcode-container mt-1"> <span><svg id="barcode"></svg></span></div>
                            <div class="sticker-inner">
                                <div class="sticker-id"><span>${number}</span></div>
                                <div class="sticker-right">
                                    <div class="sticker-name">${row[3]}</div>
                                    <div class="sticker-name"><div class="sticker-method">${row[5]}</div><div class="sticker-date">${updateDate()}</div></div>
                                    <div class="sticker-name" style="font-size: 0.5rem">${row[4]}</div>
                                </div>
                            </div>
                        </div>
                    </div>
                    `;

                    const barcodeValue = row[1];
                    console.log(barcodeValue);

                    // สร้างบาร์โค้ด
                    JsBarcode("#barcode", barcodeValue, {
                        format: "CODE39",
                        margin: 0,
                        padding: 0,
                        width: 1,
                        height: 40,
                        displayValue: false
                    });
                }
            });

            // ถ้าไม่พบข้อมูลที่ค้นหา
            if (!found) {
                sticker.innerHTML = `<p>ไม่พบข้อมูลที่ต้องการ</p>`;
            }
        })
        .catch(error => {
            console.error('Error fetching sticker data:', error);
            alert("เกิดข้อผิดพลาดในการดึงข้อมูลสติ๊กเกอร์");
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
            const number = document.getElementById('numb').textContent;
            const searchKey = document.getElementById('registernumber').textContent.trim();
            sticker.innerHTML = ''; // เคลียร์ก่อนแสดงใหม่

            let found = false; // ตัวแปรเพื่อตรวจสอบว่าพบข้อมูลหรือไม่

            data.values.forEach((row, index) => {
                if (row[0] === searchKey) {
                    found = true; // เปลี่ยนค่า found เป็น true

                    // สร้าง ID ที่ไม่ซ้ำกันสำหรับแต่ละสติ๊กเกอร์
                    const uniqueId = `barcode-${index}`;

                    sticker.innerHTML += `
                    <div class="sticker mb-2">
                        <div class="sticker-id mx-2 my-2"><span>${row[0]}</span></div>
                        <div class="sticker-right">
                            <div class="barcode-container mt-1"> <span><svg id="${uniqueId}"></svg></span></div>
                            <div class="sticker-inner">
                                <div class="sticker-id"><span>${number}</span></div>
                                <div class="sticker-right">
                                    <div class="sticker-name">${row[3]}</div>
                                    <div class="sticker-name">
                                        <div class="sticker-method">${row[5]}</div>
                                        <div class="sticker-date">${updateDate()}</div>
                                    </div>
                                    <div class="sticker-name" style="font-size: 0.5rem">${row[4]}</div>
                                </div>
                            </div>
                        </div>
                    </div>
                    `;

                    const barcodeValue = row[1];
                    console.log(barcodeValue);

                    // สร้างบาร์โค้ด
                    JsBarcode(`#${uniqueId}`, barcodeValue, {
                        format: "CODE39",
                        margin: 0,
                        padding: 0,
                        width: 1,
                        height: 40,
                        displayValue: false
                    });
                }
            });

            // ถ้าไม่พบข้อมูลที่ค้นหา
            if (!found) {
                sticker.innerHTML = `<p>ไม่พบข้อมูลที่ต้องการ</p>`;
            }
        })
        .catch(error => {
            console.error('Error fetching sticker data:', error);
            alert("เกิดข้อผิดพลาดในการดึงข้อมูลสติ๊กเกอร์");
        });
}

function printSticker() {
    closeAlert();
    const sticker = document.getElementById('sticker');
    if (sticker) {
        sticker.style.display = 'block';
    } else {
        alert("ไม่พบสติ๊กเกอร์ที่จะแสดง");
    }
}

function closeModal() {
    const sticker = document.getElementById('sticker');
    if (sticker) {
        sticker.style.display = 'none';
    }
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








function loadAllData() {
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

            data.values.forEach(row => {
                resultDiv1.innerHTML += 
                  `<tr>
                    <th scope="col" class="text-center">${row[0]}</th>
                    <td scope="col" colspan="2" class="text-center" style="font-family: sarabun;">${row[1]}</td>
                    <td scope="col" colspan="6" class="text-align-start" style="font-family: sarabun;">${row[2]}</td>
                    <td scope="col" class="text-center" style="font-family: sarabun;">${row[4]}</td>
                  </tr>`;
            });
            loadAllCount();
        })
        .catch(error => {
            console.error("Error fetching data:", error);
            alert('เกิดข้อผิดพลาดในการโหลดข้อมูล');
        });
}


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
    .then(response => {
        if (!response.ok) {
            throw new Error("Network response was not ok " + response.statusText);
        }
        return response.json();
    })
    .then(data => {
        console.log("Data added successfully", data);
        closeNewRegister();
        loadAllData(); // โหลดข้อมูลใหม่หลังจากเพิ่มข้อมูล
    })
    .catch(error => {
        console.error("Error adding data:", error);
        alert('เกิดข้อผิดพลาดในการเพิ่มข้อมูล');
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
    $('.modalregister').css('display', 'block');
}

function openSpecimen() {
    $('.modalspecimen').css('display', 'flex');
}

function closeSpecimen() {
    $('.modalspecimen').css('display', 'none');
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
    var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}?key=${apiKey}`; // แทนที่ด้วย API Key ของคุณ

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
            "Authorization": `Bearer ${access_token}`, // ใช้ OAuth token ใน header
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
        loadAllCount(); // เรียกฟังก์ชันเพื่อโหลดข้อมูลใหม่
        clearSpecimen(); // ล้างค่าใน input bar
    })
    .catch(error => {
        console.error('Error:', error);
        alert("เกิดข้อผิดพลาดในการเพิ่มข้อมูล!");
    });
}



function updateSheet(sheetName, column, data) {
    const range = `${sheetName}!${String.fromCharCode(64 + column)}2`; // ใช้คอลัมน์ที่ระบุ
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?valueInputOption=USER_ENTERED&key=${apiKey}`; // แทนที่ด้วย API Key ของคุณ

    fetch(url, {
        method: 'PUT',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(data)
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Error updating data: ' + response.statusText);
        }
        return response.json();
    })
    .then(result => {
        console.log('Data updated:', result);
    })
    .catch(error => {
        console.error('Error updating data:', error);
    });
}

function loadAllCount() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}?key=${apiKey}`;

    fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error("Network response was not ok " + response.statusText);
            }
            return response.json();
        })
        .then(data => {
            const count = data.values.length; // นับจำนวนแถวข้อมูล
            document.getElementById('dataCount').innerText = `จำนวนข้อมูลทั้งหมด: ${count}`;
        })
        .catch(error => {
            console.error("Error fetching count:", error);
            alert('เกิดข้อผิดพลาดในการโหลดจำนวนข้อมูล');
        });
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
            if (data.values && data.values.length > 0) {
                document.getElementById('register').textContent = data.values.length; // นับจำนวนแถว
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

function openSheet() {
    $(".modalsheet").css('display', 'block');
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


