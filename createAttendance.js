// createAttendance.js
const admin = require('firebase-admin');
const Excel = require('exceljs');
const moment = require('moment-timezone');
const fs = require('fs');
const os = require('os');
const path = require('path');

const SERVICE_ACCOUNT_JSON = process.env.SERVICE_ACCOUNT_JSON; // from GH secret
if(!SERVICE_ACCOUNT_JSON){
  console.error('Please set SERVICE_ACCOUNT_JSON env var (contents of service account JSON).');
  process.exit(1);
}
const serviceAccount = JSON.parse(SERVICE_ACCOUNT_JSON);

// bucket guess
const bucketName = (serviceAccount.project_id ? `${serviceAccount.project_id}.appspot.com` : process.env.STORAGE_BUCKET);

admin.initializeApp({
  credential: admin.credential.cert(serviceAccount),
  storageBucket: bucketName
});

const db = admin.firestore();
const bucket = admin.storage().bucket();

// config: set via env or defaults
const DEPT = process.env.DEPT || 'ECA';
const TZ = process.env.TZ || 'Asia/Kolkata';
const SIGNED_URL_EXPIRE_DAYS = Number(process.env.SIGNED_URL_EXPIRE_DAYS || 7);

(async function main(){
  try{
    const now = moment().tz(TZ);
    const dateStr = now.format('YYYY-MM-DD');
    const dayName = now.format('dddd'); // Monday, Tuesday, ...
    console.log('Creating attendance for', dateStr, dayName, 'dept', DEPT);

    // load timetable
    const ttSnap = await db.collection('timetables').doc(DEPT).get();
    if(!ttSnap.exists){ throw new Error(`No timetable document for dept "${DEPT}"`); }
    const ttDoc = ttSnap.data();
    const timetable = ttDoc[dayName];
    if(!Array.isArray(timetable) || timetable.length === 0){
      throw new Error(`No timetable entries for ${dayName} in dept ${DEPT}`);
    }

    // fetch students
    const studentsSnap = await db.collection('students').orderBy('rollNo', 'asc').get().catch(async err=>{
      // fallback if rollNo doesn't exist
      return await db.collection('students').get();
    });
    const students = [];
    studentsSnap.forEach(s => students.push({ id: s.id, data: s.data() }));
    console.log('Students count:', students.length);

    // prepare Excel
    const workbook = new Excel.Workbook();
    const sheet = workbook.addWorksheet('Attendance');

    const header = ['Roll', 'Name', ...timetable.map(p => `${p.subject || p.sub || 'SUB'} (${p.start}-${p.end})`)];
    sheet.addRow(header);

    // Firestore batch (commit in chunks)
    let batch = db.batch();
    let batchCount = 0;
    const BATCH_LIMIT = 400; // keep headroom under 500

    for(const st of students){
      const stData = st.data || {};
      const roll = stData.roll || stData.rollNo || stData.roll_no || stData.Roll || '';
      const name = stData.name || stData.students_name || stData.studentsName || '';
      const docId = `${dateStr}_${st.id}`;
      const docRef = db.collection('attendance_records').doc(docId);

      const rec = { date: dateStr, dept: DEPT, studentId: st.id };
      timetable.forEach((_,i)=> rec[`period${i+1}`] = 'Absent');

      batch.set(docRef, rec);
      batchCount++;
      if(batchCount >= BATCH_LIMIT){
        await batch.commit();
        batch = db.batch();
        batchCount = 0;
      }

      const rowValues = [roll, name, ...timetable.map(()=> 'Absent')];
      sheet.addRow(rowValues);
    }
    if(batchCount > 0) await batch.commit();

    // save excel to tmp
    const tmpPath = path.join(os.tmpdir(), `attendance_${DEPT}_${dateStr}.xlsx`);
    await workbook.xlsx.writeFile(tmpPath);

    // upload to storage
    const destPath = `attendance_sheets/${DEPT}_${dateStr}.xlsx`;
    await bucket.upload(tmpPath, { destination: destPath, metadata: { contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }});
    console.log('Uploaded to', destPath);

    // generate signed URL
    const file = bucket.file(destPath);
    const [signedUrl] = await file.getSignedUrl({
      action: 'read',
      expires: new Date(Date.now() + SIGNED_URL_EXPIRE_DAYS * 24*60*60*1000)
    });

    // write metadata doc
    await db.collection('attendance_sheets').doc(`${DEPT}_${dateStr}`).set({
      date: dateStr,
      dept: DEPT,
      storagePath: destPath,
      signedUrl,
      createdAt: admin.firestore.FieldValue.serverTimestamp()
    });

    console.log('Done. Signed URL (expires in', SIGNED_URL_EXPIRE_DAYS, 'days):', signedUrl);
  }catch(err){
    console.error(err);
    process.exit(1);
  }
})();
