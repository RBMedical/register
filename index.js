



function showLoading() {
    document.getElementById('loadingIcon').style.display = 'block';
}

function hideLoading() {
    document.getElementById('loadingIcon').style.display = 'none';
}

function showLoadingSearch() {
    document.getElementById('loadingIconSearch').style.display = 'block';
}
function hideLoadingSearch() {
    document.getElementById('loadingIconSearch').style.display = 'none';
}

const apiKey = 'AIzaSyBxpc9yQw66EWefkTU9O7ITp-bvN72eESg';
const spreadsheetId = '1_aUWV9uDvVn_WBs25ZsHtVLilUYB9iNP87yadjSbHsw';
const rangesheet1 = 'data!A3:ZZ';
const rangesheet2 = 'program!A2:ZZ';
const rangesheet3 = 'register!A2:ZZ';
const rangesheet4 = 'register!A2:A';
const rangesheet5 = 'sticker!A2:ZZ';
const rangesheet6 = 'specimencount!A2:ZZ';
const rangesheet7 = 'specimencount!A:A';
const rangesheet8 = 'program!A2:ZZ';
const rangesheet9 = 'sticker!A2:ZZ';
const rangesheet10 = 'register!A2:ZZ';
const rangesheet11 = 'register!K2:KK';
const rangesheet12 = 'specimencount!B1:EE';

const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');

// กำหนดไฟล์ JSON Key ที่ได้จาก Service Account
const keyPath = path.join(__dirname, 'your-service-account-key.json');
const key = JSON.parse(fs.readFileSync(keyPath));
const PRIVATE_KEY = `-----BEGIN PRIVATE KEY-----\nMIIEvgIBADANBgkqhkiG9w0BAQEFAASCBKgwggSkAgEAAoIBAQDZwbpuppgpynmW\nu/ImSDad+dv1KiWWCNBE/
                     Uib5iLNY6+Mhn8JT7pr25i52OZjHF8paZfBJkozWCBY\n0rP1RKM1wfYz2mUwDQicRubtac/h8B9kA/JMmpx+OeR6lPtGowjVn/ha9JS/4fwd\
                     n51yZxQINkg1mE2bRMOuv1XvXW/71p8GcWw3MiUKAIH1Es0LcN22Dsj8lrtowy4ze\n6il3gTz7tWz0yOhSW0CYk6TpBAsSDjlRRh00hRs49M37Om9AkiVxQTPt1MPHnDYI\
                     ni9SSrOtIx/zIgTlCoVFTCHZKHw810G58AgahSRxatxI7fzy91fy9Ub2WYfCXDAog\nnoutVxRrAgMBAAECggEAD5H+Vym69dRs/XeCso5CULYMR/6Tmz1JF+dc33NFGzBc\
                     n3cYpY9lbSDNDbrKELFaoghXZsu1rRW2W1OyuaN2K2aEWjis4Rgdm8+m1jgjqKbEm\nRXWeYZZ1yMH+xNPnSoWfPMsuxa4sZRk73xWE6QsNtDoa+PWh/SSrpzIZvbkg+DnW\
                     nafuZBVseEwqGuYcZf+B65ePUegtX8/FlXvJmfI470vO2cX7piV2icr6YgU9SQD2r\nHFh/u3b8k1qBIV9UxEUX9HjOTXGG1t0/2/AD55E5+9ztWOPwtr53BYiTLgpUpC+x\
                     nJ4vLCE+xtThuiecgvNwmPaPL9pm09W3eVZIVIXcV4QKBgQD9rgDwxDoMxFKsXBYP\nhrjcZZlm4wjZJmKyngzu4oYnTUXSjvGee07npsq2l3+5/nJaZzsKsRp6kPBE9pN0\
                     nhVkRgZtVRXgZ5jtxLgaM5wm9vJG3vA0ApmB25MQtndCbml73bWQ8qbUUq09Lrdj7\nas2aluh56vjnFH1ZNQWS7QcJYQKBgQDbv5w4dYXdfhru4YhhMmP4FGd7AknEfAPF\
                     nQyecWX5T/L7cLUpwxTPKDmn0Xv9+CIz/bG7J/S9aKPd1MFpFlFFhmXotqQ3d4gHB\ne9Dpk8o1QEVNoUg/3xrT0hSsZKi6Y9FHHK5PY4v8y/Phiz9f2LqFDE03oR0/B1gc\
                     n0/8ib751SwKBgQC2nqwIp4qOpEpL0GMFPFwaNX3QZoJ5KLwGj+cJlcMzydoI8WSZ\nTXWJKDZoafnGIJmb4RLM6KACOhLt4oBWcqSjCKWVJlSGeIq0OIj4qF4H3BceqN7H\
                     nZ/6ruJZNrH1/dwsEnhh530X/oi+McJNysvleX2LuWaxjVgnCzXu8wKu/IQKBgQCD\nfVkOE4yBZ4bYL826UzusYxE0cr8POiHLdI6MKKTFvrO57cPgTK/blNpjpkB8+sLb\
                     nx9dXOA+QhHjl/4PUpJY5r2uDTOgGP8lLLDpquctCJ+4QMJSZ23cjDk7ehPDNbxL3\n2TqYOHm4T5Xj/L10LawWFrFRuy9T2qInxdahlXnClwKBgBdFN5Yf8dw+rGfvUU0Q\
                     nk7bzrls36YpMTsuWs6zm9YLsy6wgUTOveaMNElH6iadT7HN4jrZt7huhkhANFzLm\n6+dkiUFCtu9ZxTV67VFnDbaAyMOhU/WdUCOibcnbRY48Vfzh0bIXeXMmSGyZYr4G\
                     nfQglRejjUtp9LTkhLvx6/3RN\n-----END PRIVATE KEY-----\n`; 

const SCOPES = 'https://www.googleapis.com/auth/spreadsheets';


function generateJWT() {
    const header = { alg: 'RS256', typ: 'JWT' };
    const now = Math.floor(Date.now() / 1000);
    const payload = {
        iss: CLIENT_EMAIL,
        scope: SCOPES,
        aud: 'https://oauth2.googleapis.com/token',
        exp: now + 3600,
        iat: now
    };
    const sHeader = JSON.stringify(header);
    const sPayload = JSON.stringify(payload);
    return KJUR.jws.JWS.sign(null, sHeader, sPayload, PRIVATE_KEY);
}

function fetchAccessToken(callback) {
    const jwt = generateJWT();
    fetch('https://oauth2.googleapis.com/token', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
            grant_type: 'urn:ietf:params:oauth:grant-type:jwt-bearer',
            assertion: jwt,
        })
    })
    .then(response => response.json())
    .then(data => {
        callback(data.access_token);
    })
    .catch(error => {
        console.error('Error fetching access token:', error);
        Swal.fire({
            icon: 'error',
            title: 'เกิดข้อผิดพลาด',
            text: 'ไม่สามารถดึง Access Token ได้!'
        });
    });
}

function addRegistrationData() {
    const regisData = [
        document.getElementById('numb').textContent.trim(),
        document.getElementById('registernumber').textContent.trim(),
        document.getElementById('name').textContent.trim(),
        document.getElementById('idcard').textContent.trim(),
        document.getElementById('card').textContent.trim(),
        document.getElementById('depart').textContent.trim(),
        document.getElementById('age').textContent.trim(),
        document.getElementById('birthday').textContent.trim(),
        document.getElementById('program').textContent.trim(),
        document.getElementById('datetime').textContent.trim(),
        document.getElementById('desc').textContent.trim()
    ];

    const data = { values: [regisData] };

    fetchAccessToken(function(accessToken) {
        const checkUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet3}?key=${apiKey}`;
        
        fetch(checkUrl, {
            method: 'GET',
            headers: { 'Authorization': `Bearer ${accessToken}` }
        })
        .then(checkResponse => checkResponse.json())
        .then(sheetData => {
            const rows = sheetData.values || [];
            const isDuplicate = rows.some(row => row[3] === regisData[3]);

            if (isDuplicate) {
                Swal.fire({ icon: 'error', title: 'Oops...', text: 'ID นี้ลงทะเบียนแล้ว!' });
                return;
            }

            const postUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet3}:append?valueInputOption=USER_ENTERED&key=${apiKey}`;
            
            fetch(postUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${accessToken}`
                },
                body: JSON.stringify(data)
            })
            .then(postResponse => {
                if (postResponse.ok) {
                    Swal.fire({
                        position: "center",
                        icon: "success",
                        title: "ลงทะเบียนสำเร็จ",
                        showConfirmButton: false,
                        timer: 1500
                    });
                     addRegistrationDataInner();
                } else {
                    throw new Error('Error in adding data');
                }
            })
            .catch(error => {
                Swal.fire({
                    icon: 'error',
                    title: 'เกิดข้อผิดพลาด',
                    text: 'ไม่สามารถเพิ่มข้อมูลได้!'
                });
                console.error('Error:', error);
            });
        })
        .catch(error => {
            Swal.fire({
                icon: 'error',
                title: 'เกิดข้อผิดพลาด',
                text: 'ไม่สามารถตรวจสอบข้อมูลได้!'
            });
            console.error('Error:', error);
        });
    });
}

function addRegistrationDataInner() {
    const innerData = [
        document.getElementById('datetime').textContent.trim(),
        document.getElementById('registernumber').textContent.trim(),
        document.getElementById('name').textContent.trim(),
        "10",
        "1ลงทะเบียน"
    ];

    const data = { values: [innerData] };

    fetchAccessToken()
        .then(accessToken => {
            const postUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}:append?valueInputOption=USER_ENTERED&key=${apiKey}`;
            return fetch(postUrl, {
                method: 'POST',
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': `Bearer ${accessToken}`
                },
                body: JSON.stringify(data)
            });
        })
        .then(postResponse => {
            if (!postResponse.ok) {
                throw new Error('Error in adding specimen count data');
            }
            console.log("Data added to specimen count successfully");
        })
        .catch(error => {
            Swal.fire({ icon: 'error', title: 'เกิดข้อผิดพลาด', text: 'ไม่สามารถเพิ่มข้อมูล specimen ได้!' });
            console.error('Error:', error);
        });
}

function addNewData() {
    var newid = document.getElementById('newid').value.trim();
    var newname = document.getElementById('newname').value.trim();
    var newidcard = document.getElementById('newidcard').value.trim();
    var birthdate = document.getElementById('birthdate').value.trim();
    var newcard = document.getElementById('newcard').value.trim();
    var newdepart = document.getElementById('newdepart').value.trim();
    var newage = document.getElementById('newage').textContent.trim();
    var newprogram = document.getElementById('newprogram').value.trim();

    if (!newid || !newname || !newidcard || !birthdate || !newcard || !newdepart || !newage || !newprogram) {
        alert('กรุณากรอกข้อมูลให้ครบถ้วน');
        return;
    }

    var newRow = [newid, newname, newidcard, newcard, newdepart, newage, birthdate, newprogram];

    fetchAccessToken()
        .then(accessToken => {
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}:append?valueInputOption=USER_ENTERED&key=${apiKey}`;
            const body = { values: [newRow] };

            return fetch(url, {
                method: "POST",
                headers: {
                    'Content-Type': 'application/json',
                    'Authorization': 'Bearer ' + accessToken,
                },
                body: JSON.stringify(body)
            });
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
function getNextNumber() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet4}?key=${apiKey}`;
    
    return fetchAccessToken()
        .then(accessToken => {
            return fetch(url, {
                headers: { 'Authorization': `Bearer ${accessToken}` }
            });
        })
        .then(response => {
            if (!response.ok) {
                throw new Error("Network response was not ok " + response.statusText);
            }
            return response.json();
        })
        .then(data => {
            const values = data.values;
            return values ? values.length + 1 : 1;
        })
        .catch(error => {
            console.error("Error getting next number:", error);
        });
}



function displayNextNumber() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet4}?key=${apiKey}`;
     const accessToken = fetchAccessToken();


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
                var newNumber = document.getElementById('numb');
                newNumber.innerText = parseInt(lastNumber) + 1;
            } else {
                var newNumber = document.getElementById('numb');
                newNumber.innerText = 1;
            }
        })
        .catch(error => {
            console.error('Error fetching data:', error);
            throw error; // ส่งต่อ error ไปยัง function ที่เรียกใช้
        });
}




function searchData() {
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
            const programName = document.getElementById('program').textContent;
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
 <th scope="col" class="col-1 text-center">${row[0]}</th>
 <td scope="col" class="col-2 text-center" style="font-family: sarabun;">${row[1]}</td>
 <td scope="col" class="col-4" style="font-family: sarabun;">${row[2]}</td>
 <td scope="col" class="col-2 text-center" style="font-family: sarabun;">${row[8]}</td>
 <td scope="col" class="col-3 text-center" style="font-family: sarabun;">${row[9]}</td>
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
                     <div class="sticker-id mx-2 my-2 text-center"><span>${row[0]}</span></div>
                     <div class="sticker-right">
                         <div class="barcode-container mt-1">
                         <div class="mt-2 ms-2"><span class="libre-barcode-39-extended-regular">${row[1]}</span></div>
                         </div>
                         <div class="sticker-inner">
                             <div class="sticker-id fw-bold" style="font-size: 0.75rem;"><span>${number}</span></div>
                             <div class="sticker-right">
                                 <div class="sticker-name ms-2">${row[3]}</div>
                                 <div class="sticker-name">
                                     <div class="sticker-method ms-2">${row[5]}</div>
                                     <div class="sticker-date mt-1">${updateDate()}</div>
                                 </div>
                                 <div class="sticker-name ms-4" style="font-size: 0.5rem">${row[4]}</div>
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
    
    $(".modalrecieve").css('display', 'block');
}

function closeSticker() {
    $(".modalrecieve").css('display', 'none');
}

function printSpecimen() {
    printResult();
    closeSticker();
    clearPage();
    displayNextNumber();

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



function loadAllCount() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}?key=${apiKey}`;
    
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
            let a = 0, b = 0, c = 0, d = 0, e = 0, f = 0, g = 0, h = 0, i = 0, j = 0, k = 0, l = 0;

            data.values.forEach(row => {
                const barcodemethod = row[3]; // ตรวจสอบข้อมูลใน index 3

                // นับจำนวนการเกิดของแต่ละ `barcodemethod`
                switch (barcodemethod) {
                    case "10":
                        l++;
                        break;
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


            document.getElementById('PE').textContent = a;
            document.getElementById('EDTA').textContent = b;
            document.getElementById('urine').textContent = c;
            document.getElementById('xray').textContent = d;
            document.getElementById('ekg').textContent = e;
            document.getElementById('ear').textContent = f;
            document.getElementById('lung').textContent = g;
            document.getElementById('eye').textContent = h;
            document.getElementById('muscle').textContent = i;
            document.getElementById('naf').textContent = j;
            document.getElementById('clot').textContent = k;
            document.getElementById('register').textContent = l;

            clearSpecimen();
            loadDataTable();
        })
        .catch(error => {
            console.error('Error fetching data:', error);
        });

}


function loadAllCountOpen() {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet6}?key=${apiKey}`;
   

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
            let a = 0, b = 0, c = 0, d = 0, e = 0, f = 0, g = 0, h = 0, i = 0, j = 0, k = 0, l = 0;

            data.values.forEach(row => {
                const barcodemethod = row[3]; // ตรวจสอบข้อมูลใน index 3

                // นับจำนวนการเกิดของแต่ละ `barcodemethod`
                switch (barcodemethod) {
                    case "10":
                        l++;
                        break;
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


            document.getElementById('PE').textContent = a;
            document.getElementById('EDTA').textContent = b;
            document.getElementById('urine').textContent = c;
            document.getElementById('xray').textContent = d;
            document.getElementById('ekg').textContent = e;
            document.getElementById('ear').textContent = f;
            document.getElementById('lung').textContent = g;
            document.getElementById('eye').textContent = h;
            document.getElementById('muscle').textContent = i;
            document.getElementById('naf').textContent = j;
            document.getElementById('clot').textContent = k;
            document.getElementById('register').textContent = l;

        })
        .catch(error => {
            console.error('Error fetching data:', error);
        });

}






function loadDataTable() {

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

        
            const sortedData = data.values.sort((a, b) => {
                const valueA = parseInt(a[0], 10); // แปลงค่าเป็นตัวเลข
                const valueB = parseInt(b[0], 10);
                return valueB - valueA; // เรียงจากมากไปน้อย
            });

            // แสดงข้อมูลใน resultDiv1
            sortedData.forEach(row => {
                if (row[4] !== "1ลงทะเบียน") {
                resultDiv1.innerHTML +=
                    `<tr>
                 <th scope="row" class="col-3 text-center">${row[0]}</th>
                 <td scope="col"  class="col-2 text-center" style="font-family: sarabun;">${row[1] || 'N/A'}</td>
                 <td scope="col" colspan="6" class="text-align-start" style="font-family: sarabun;">${row[2] || 'N/A'}</td>
                 <td scope="col" class="text-center" style="font-family: sarabun;">${row[4] || 'N/A'}</td>
               </tr>`;
                }
            });

            // เรียกใช้ loadAllCount() หลังจากแสดงผลข้อมูล

        })
        .catch(error => {
            console.error("Error fetching data:", error);
            alert('เกิดข้อผิดพลาดในการโหลดข้อมูล');
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







function sendBarcode() {

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
                setTimeout(() => {
                    addRegistData();
                }, 1000);
            } else {
                alert('ไม่พบ ID นี้ในระบบ');
            }
        })
        .catch(error => {
            console.error('Error in fetching barcode data:', error);
            alert('เกิดข้อผิดพลาดในการค้นหาข้อมูล');
        });

}




function addRegistData() {
  



    var barinput = document.getElementById('inputbar').value.trim();
    var barcodenewid = document.getElementById('barregisterid').textContent.trim();
    var barcodename = document.getElementById('barname').textContent.trim();
    var inputdate = document.getElementById('datetime').textContent;
    var inputmethodbar = barinput.slice(-2); // ดึง 2 ตัวท้ายของบาร์โค้ด 
    console.log(inputmethodbar);
    var barinputmethod = inputmethodbar;
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



    var data = {
        values: [[inputdate, barcodenewid, barcodename, barinputmethod, specimen]]
    };
    console.log(data);
    const accessToken = fetchAccessToken();
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
            loadAllData();

        })
        .catch(error => {
            console.error('Error:', error);
            alert("เกิดข้อผิดพลาดในการเพิ่มข้อมูล!");
        });


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

            // เรียงลำดับข้อมูลจากคอลัมน์ [0] จากมากไปน้อย
            const sortedData = data.values.sort((a, b) => {
                const valueA = parseInt(a[0], 10); // แปลงค่าเป็นตัวเลข
                const valueB = parseInt(b[0], 10);
                return valueB - valueA; // เรียงจากมากไปน้อย
            });

            // แสดงข้อมูลใน resultDiv1
            sortedData.forEach(row => {
                if (row[4] !== "1ลงทะเบียน") {
                    resultDiv1.innerHTML +=
                        `<tr>
                 <th scope="row" class="col-3 text-center">${row[0]}</th>
                 <td scope="col" class="col-2 text-center" style="font-family: sarabun;">${row[1] || 'N/A'}</td>
                 <td scope="col" colspan="6" class="text-align-start" style="font-family: sarabun;">${row[2] || 'N/A'}</td>
                 <td scope="col" class="text-center" style="font-family: sarabun;">${row[4] || 'N/A'}</td>
               </tr>`;
                }
            });

            // เรียกใช้ loadAllCount() หลังจากแสดงผลข้อมูล
            loadAllCount();
        })
        .catch(error => {
            console.error("Error fetching data:", error);
            alert('เกิดข้อผิดพลาดในการโหลดข้อมูล');
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

function clearSpecimen() {

    var barcode = document.getElementById('inputbar');
    var bid = document.getElementById('barregisterid');
    var bname = document.getElementById('barname');
    barcode.value = '';
    bid.innerText = '';
    bname.innerText = '';
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
function closeRegist() {
    clearRegisterPage();
    $(".modalregister").css('display', 'none');
}
function closeSheet() {
    $(".modalsheet").css('display', 'none');
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
                    const n3 = document.getElementById('desc');
                    n2.innerText = n1;
                    n3.innerText = "เพิ่มรายชื่อ";

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
function loadAllRecordsSearch() {


const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}?key=${apiKey}`;
 
fetch(url)
    .then(response => {
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        return response.json();
    })
    .then(data => {
        const resultDiv1 = document.getElementById('searchdataload');

        resultDiv1.innerHTML = ''; // เคลียร์ผลลัพธ์ก่อนแสดงใหม่

        if (data.values && data.values.length > 0) { // ตรวจสอบว่ามีข้อมูลใน data.values
            data.values.forEach(row => {
                resultDiv1.innerHTML += `
<tr onclick="copyId" style="cursor: pointer;">

          <th class="column-5 text-center">${row[0]}</th>
          <td class="column-6">${row[1]}</th>
          <td class="column-5 text-center">${row[2]}</th>
          <td class="column-5 text-center">${row[4]}</th>
        
</tr>
`;
 hideLoadingSearch();

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

function openSearch() {
    $(".modalsearch").css('display', 'block');
     clearPage();
  
}

function closeSearch() {
    $(".modalsearch").css('display', 'none');
}


function searchDataSearch(){
   
      const searchId = document.getElementById('datasearchid').value.trim(); // ดึงค่าจาก input และลบช่องว่าง

     const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}?key=${apiKey}`;

    fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
             var resultDiv1 = document.getElementById('searchdataload');
                resultDiv1.innerHTML = '';
            if (data.values) {
                data.values.forEach(row => {
                    if (row[2] === searchId){
                     
                          resultDiv1.innerHTML += `
<tr onclick="copyId"  style="cursor: pointer;">

          <th class="column-5 text-center">${row[0]}</th>
          <td class="column-6">${row[1]}</th>
          <td id="copyid" class="column-5 text-center">${row[2]}</th>
          <td class="column-5 text-center">${row[4]}</th>
        
</tr>
`;

               } });
                 
        } else {
            resultDiv1.innerHTML = `<tr><td colspan="8" class="text-center">ไม่พบข้อมูล</td></tr>`;
        }
    })
    .catch(error => {
        console.error('Error loading all records:', error);
        alert("เกิดข้อผิดพลาดในการดึงข้อมูล");
    });
  
}

function copyId(element) {
    // element คือแถวที่ถูกคลิก
    const idcard = element.querySelector(".idCardCell").textContent.trim();
    const copyend = document.getElementById('idcard');

    if (copyend) {
        copyend.textContent = idcard; 
    }

    closeSearch();
     setTimeout(() => {
                       searchDataFromId();
                                    }, 5000);
   
}


function filterRecords() {
    const searchValue = document.getElementById('datasearchname').value.toLowerCase();
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet1}?key=${apiKey}`;

    fetch(url)
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
            const resultDiv1 = document.getElementById('searchdataload');
            resultDiv1.innerHTML = ''; // เคลียร์ข้อมูลเก่าออก

            if (data.values && data.values.length > 0) {
                data.values.forEach(row => {
                    // กรองเฉพาะแถวที่คอลัมน์ที่สอง (row[1]) มีข้อความที่คล้ายกับ searchValue
                    if (row[1] && row[1].toLowerCase().includes(searchValue)) {
                        resultDiv1.innerHTML += `
                            <tr onclick="copyId(this)" style="cursor: pointer;">
                                <th class="column-5 text-center">${row[0]}</th>
                                <td class="column-6">${row[1]}</td>
                                <td class="idCardCell column-5 text-center">${row[2]}</td>
                                <td class="column-5 text-center">${row[4]}</td>
                            </tr>
                        `;
                    }
                });

                // หากไม่พบข้อมูลที่ตรงกับการค้นหา
                if (resultDiv1.innerHTML === '') {
                    resultDiv1.innerHTML = `<tr><td colspan="4" class="text-center">ไม่พบข้อมูลที่ตรงกับการค้นหา</td></tr>`;
                }
            } else {
                resultDiv1.innerHTML = `<tr><td colspan="4" class="text-center">ไม่พบข้อมูล</td></tr>`;
            }
        })
        .catch(error => {
            console.error('Error loading records:', error);
            alert("เกิดข้อผิดพลาดในการดึงข้อมูล");
        });
}

function deleteAllFilter(){
     loadAllRecordsSearch();
     var search1 = document.getElementById('datasearchid');
     var search2 = document.getElementById('datasearchname');
    search1.value = '';
    search2.value = '';
}

async function fetchData() {
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${rangesheet12}?key=${apiKey}`;
            const response = await fetch(url);
            const data = await response.json();
            
            const formattedData = [];
            const headers = data.values[0];
            for (let i = 1; i < data.values.length; i++) {
                const row = {};
                headers.forEach((header, index) => {
                    row[header] = data.values[i][index];
                });
                formattedData.push(row);
            }
            return formattedData;
        }

        fetchData().then(data => {
            $("#output").pivotUI(data, {
                rows: ["Register ID", "ชื่อ นามสกุล"], 
                cols: ["Specimen"],
                vals: ["Specimen"],
                aggregatorName: "Count",
                rendererName: "Table"
            });
        });

window.onload = function () {
  loadAllRecords();
  getNextNumber(); 
  updateDateTime();
  loadDataTable();
  loadAllCount();
 
}


