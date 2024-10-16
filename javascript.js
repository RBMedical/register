// ข้อมูล Client
const apiKey = 'AIzaSyDvWPO5V4Iin8OpQoq0otZNmcicWayo-6E';
const spreadsheetId = '1_aUWV9uDvVn_WBs25ZsHtVLilUYB9iNP87yadjSbHsw';
const rangesheet1 = 'data!A2:ZZ'; 
const rangesheet2 = 'program!A2:ZZ';
const rangesheet3 = 'register!A2:ZZ';
const rangesheet4 = 'register!A:A';
const rangesheet5 = 'sticker!A2:zz';
const rangesheet6 = 'specimencount!A2:zz';

const clientId = "168121174551-p6j0heikm2aajscj33ngja68s36t35nr.apps.googleusercontent.com";
const clientSecret = "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";
const redirectUri = "https://rbmedical.github.io/register";

// ข้อมูล Token (ที่คุณให้มา)
const tokenData = {
    access_token: "ya29.a0AcM612zUy-IOhLhuQcXCfQN0-uv89Is0a4ZqiRKw8UtsK7F2_KCtlCWRMASi6B94U5Iv8pGmLQvb3sgOUytYzqV9ogkYcnYAjJ7aTjTPPQA7S1kz3mRhOho45jOibtp2iAVaB0BqzRuhg9ow1vMVHT7VhKDBFg0Vo6_QvfaxaCgYKAbUSARESFQHGX2Mi_hQtl_Xk8pxh3zePOimZbg0175",
    expires_in: 3599,
    refresh_token: "1//05NVsEHUGIgFnCgYIARAAGAUSNwF-L9Irbtd0caSfh03jfuqH1ydgufcYl69gawRKI1fcBYPm5ku3mLGEBP0BUL1JQXBKv2Q6JUE",
    scope: "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/spreadsheets",
    token_type: "Bearer"
};

// ฟังก์ชันที่เรียกเมื่อหน้าโหลด
window.onload = function() {
    // เก็บ access token และ refresh token ใน sessionStorage
    sessionStorage.setItem("access_token", tokenData.access_token);
    sessionStorage.setItem("refresh_token", tokenData.refresh_token);
    console.log("Access token:", tokenData.access_token);
    console.log("Refresh token:", tokenData.refresh_token);

    // ตรวจสอบว่า access token หมดอายุหรือไม่
    setInterval(checkAndRefreshToken, (tokenData.expires_in - 60) * 1000); // รีเฟรชก่อนหมดอายุ 1 นาที
};

// ฟังก์ชันสำหรับรีเฟรช access token
function refreshAccessToken() {
    const refreshToken = sessionStorage.getItem("refresh_token");

    const url = "https://oauth2.googleapis.com/token";

    const params = new URLSearchParams();
    params.append("grant_type", "refresh_token");
    params.append("client_id", clientId);
    params.append("client_secret", clientSecret);
    params.append("refresh_token", refreshToken);

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
    });
}

// ฟังก์ชันเพื่อตรวจสอบและรีเฟรช token เมื่อจำเป็น
function checkAndRefreshToken() {
    const accessToken = sessionStorage.getItem("access_token");

    console.log("Checking if access token needs refresh...");
    refreshAccessToken(); // เรียกฟังก์ชันรีเฟรชทุกครั้ง (สามารถปรับให้ตรวจสอบได้)
}




function searchData() {
    const searchKey = document.getElementById('searchKey').value.trim(); // ดึงค่าจาก input และลบช่องว่าง
    const accessToken = sessionStorage.getItem("access_token"); // ดึง access token จาก sessionStorage
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}?access_token=${accessToken}`;

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
    const accessToken = sessionStorage.getItem("access_token"); // ดึง access token จาก sessionStorage
    const url1 = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet2}?access_token=${accessToken}`;

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


function loadAllRecords() {
    checkAndRefreshToken();
    const accessToken = sessionStorage.getItem("access_token"); // ดึง access token จาก sessionStorage
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet3}?access_token=${accessToken}`;

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


window.onload = function(){
  loadAllRecords();
//    displayNextNumber();
    updateDateTime();
    loadAllData();
}
    







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
    
    checkAndRefreshToken(); // ตรวจสอบและรีเฟรช token
    
    // รอให้ access_token ถูกอัปเดตใน sessionStorage ก่อนทำการเพิ่มข้อมูล
    const accessToken = sessionStorage.getItem("access_token");
    var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet3}:append?valueInputOption=USER_ENTERED&key=${apiKey}`; // แทนที่ด้วย API Key ของคุณ

    // ส่งข้อมูลไปยัง Google Sheets
    fetch(url, {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + accessToken // ใช้ access_token ที่ดึงจาก sessionStorage
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
        loadAllRecords(); // เรียกใช้ฟังก์ชันเพื่อโหลดข้อมูลใหม่
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
                                <span><svg id="${uniqueId}"></svg></span>
                            </div>
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





function loadAllData() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}?key=${apiKey}`;
    checkAndRefreshToken();
    
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
}




function addNewData(access_token) {
    // ดึงข้อมูลจาก input elements
    var newid = document.getElementById('newid').value.trim();
    var newname = document.getElementById('newname').value.trim();
    var newidcard = document.getElementById('newidcard').value.trim();
    var birthdate = document.getElementById('birthdate').value.trim();
    var newage = document.getElementById('newage').textContent.trim(); // ใช้ innerText หรือ textContent ให้เหมาะสมกับ HTML
    var newprogram = document.getElementById('newprogram').value.trim();

    // ตรวจสอบว่าข้อมูลทั้งหมดถูกกรอก
    if (!newid || !newname || !newidcard || !birthdate || !newage || !newprogram) {
        alert('กรุณากรอกข้อมูลให้ครบถ้วน');
        return; // ออกจากฟังก์ชันถ้ามีข้อมูลไม่ครบ
    }

    var newRow = [newid, newname, newidcard, birthdate, newage, newprogram];
    checkAndRefreshToken(); // ตรวจสอบและรีเฟรช OAuth token

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
        closeNewRegister(); // ปิด modal หรือ form ที่เพิ่มข้อมูล
        loadAllData(); // โหลดข้อมูลใหม่หลังจากเพิ่มข้อมูล
        Swal.fire({
            position: 'center',
            icon: 'success',
            title: 'เพิ่มข้อมูลสำเร็จ',
            showConfirmButton: false,
            timer: 1500
        });
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

    var barcode = document.getElementById('inputbar').value.trim();
    var barcodeid = barcode.substring(0, 8); // เอา 8 ตัวแรกของบาร์โค้ดมา

    // URL สำหรับ Google Sheets API
    var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}?key=${apiKey}`;

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
            var foundRecord = records.find(record => record[0] === barcodeid); // ค้นหาข้อมูลที่ตรงกับ barcodeid

            var baridElement = document.getElementById('barregisterid');
            var barnameElement = document.getElementById('barname');

            // เคลียร์ค่าจาก elements
            baridElement.textContent = '';
            barnameElement.textContent = '';

            if (foundRecord) {
                baridElement.textContent = foundRecord[0]; // barid
                barnameElement.textContent = foundRecord[1]; // barname

                var id = baridElement.textContent;
                var name = barnameElement.textContent;
                console.log(id, name);
                addRegistData(); // ฟังก์ชันที่ใช้เพิ่มข้อมูล
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
            specimen = "EDTA";
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
        case '20':
            specimen = "naf";
            break;
        case '21':
            specimen = "Clot";
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
    checkAndRefreshToken();

    var url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}:append?valueInputOption=USER_ENTERED&key=${apiKey}`;

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

        // หลังจากเพิ่มข้อมูลเสร็จแล้ว เรียกฟังก์ชัน updateSheet เพื่อตรวจสอบ barinputmethod และอัปเดตข้อมูลในชีต 'data'
        updateDataSheet(barcodenewid, barinputmethod);
    })
    .catch(error => {
        console.error('Error:', error);
        alert("เกิดข้อผิดพลาดในการเพิ่มข้อมูล!");
    });
}

// ฟังก์ชันที่อัปเดตข้อมูลใน sheet 'data'
function updateDataSheet(barcodenewid, barinputmethod) {
    // ระบุคอลัมน์ที่จะอัปเดตตามเงื่อนไขของ barinputmethod
    let columnToUpdate;
    switch (barinputmethod) {
        case '11':
            columnToUpdate = 7;  // คอลัมน์ [7]
            break;
        case '12':
            columnToUpdate = 8;  // คอลัมน์ [8]
            break;
        case '13':
            columnToUpdate = 11; // คอลัมน์ [11]
            break;
        case '14':
            columnToUpdate = 12; // คอลัมน์ [12]
            break;
        case '15':
            columnToUpdate = 13; // คอลัมน์ [13]
            break;
        case '16':
            columnToUpdate = 14; // คอลัมน์ [14]
            break;
        case '17':
            columnToUpdate = 15; // คอลัมน์ [15]
            break;
        case '18':
            columnToUpdate = 16; // คอลัมน์ [16]
            break;
        case '19':
            columnToUpdate = 17; // คอลัมน์ [17]
            break;
        case '20':
            columnToUpdate = 9;  // คอลัมน์ [9]
            break;
        case '21':
            columnToUpdate = 10; // คอลัมน์ [10]
            break;
        default:
            console.log('ไม่พบ method ที่ต้องการอัปเดต');
            return; // ออกจากฟังก์ชันหากไม่มีค่าใน switch
    }

    // สร้าง object เพื่ออัปเดตค่าในคอลัมน์ที่กำหนด
    var dataToUpdate = {
        values: [["x"]]
    };

    // URL สำหรับอัปเดตข้อมูลในชีต 'data' ตาม barcodenewid และคอลัมน์ที่ระบุ
    const range = `data!${String.fromCharCode(64 + columnToUpdate)}${barcodenewid}`;
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${range}?valueInputOption=USER_ENTERED&key=${apiKey}`;

    // ส่งคำขอไปยัง Google Sheets เพื่ออัปเดตข้อมูล
    fetch(url, {
        method: 'PUT',
        headers: {
            'Content-Type': 'application/json',
            "Authorization": `Bearer ${access_token}`, // ใช้ OAuth token ใน header
        },
        body: JSON.stringify(dataToUpdate)
    })
    .then(response => {
        if (!response.ok) {
            throw new Error('Network response was not ok: ' + response.statusText);
        }
        return response.json();
    })
    .then(data => {
        console.log('Data updated successfully in sheet "data"', data);
    })
    .catch(error => {
        console.error('Error updating data in sheet "data":', error);
    });
}


function loadAllCount() {
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

    loadRegister(); 
    clearSpecimen();  
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

function clearSpecimen(){
    var barcode = document.getElementById('inputbar');
    barcode.value = '';
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


