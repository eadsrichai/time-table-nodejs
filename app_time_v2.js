const express = require('express');
const fs = require('fs');
const csv = require('csv-parser');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 3000;

// 1. ฟังก์ชันโหลดข้อมูล
async function readCSV(filePath) {
    const results = [];
    return new Promise((resolve) => {
        if (!fs.existsSync(filePath)) return resolve([]);
        fs.createReadStream(filePath)
            .pipe(csv())
            .on('data', (data) => {
                const cleanData = {};
                Object.keys(data).forEach(key => { 
                    const cleanKey = key.trim().replace(/^\ufeff/, '');
                    cleanData[cleanKey] = data[key] ? data[key].trim() : ''; 
                });
                results.push(cleanData);
            })
            .on('end', () => resolve(results))
            .on('error', (err) => resolve([]));
    });
}

// 2. Logic จัดตาราง
function solveTimetable(teachers, rooms, subjects, timeslots, registers, teachAssignments, groups) {
    let timetable = [];
    let unassignedLog = [];
    
    let groupDailyLoad = {}; 
    let teacherDailyLoad = {};
    let roomUsageCount = {};
    
    rooms.forEach(r => {
        const rId = r.room_id || r['room id'];
        roomUsageCount[rId] = 0;
    });

    const getLoad = (dailyMap, id, day) => {
        if (!dailyMap[id]) dailyMap[id] = { Mon:0, Tue:0, Wed:0, Thu:0, Fri:0 };
        return dailyMap[id][day];
    };

    const addLoad = (dailyMap, id, day, amount) => {
        if (!dailyMap[id]) dailyMap[id] = { Mon:0, Tue:0, Wed:0, Thu:0, Fri:0 };
        dailyMap[id][day] += amount;
    };

    const isSlotFree = (tSlotId, teacherId, groupId, roomId) => {
        if (teacherId) {
            const teacherBusy = timetable.find(t => t.timeslot_id === tSlotId && t.teacher_id === teacherId);
            if (teacherBusy) return false;
        }
        const groupBusy = timetable.find(t => t.timeslot_id === tSlotId && t.group_id === groupId);
        if (groupBusy) return false;
        if (roomId) {
            const roomBusy = timetable.find(t => t.timeslot_id === tSlotId && t.room_id === roomId);
            if (roomBusy) return false;
        }
        return true;
    };

    const validRegisters = registers.filter(r => r.subject_id || r['subject id']);

    const sortedRegisters = [...validRegisters].sort((a, b) => {
        const sIdA = a.subject_id || a['subject id'];
        const sIdB = b.subject_id || b['subject id'];
        const subA = subjects.find(s => (s.subject_id || s['subject id']) === sIdA);
        const subB = subjects.find(s => (s.subject_id || s['subject id']) === sIdB);
        const totalA = parseInt(subA?.theory||0) + parseInt(subA?.practice||0);
        const totalB = parseInt(subB?.theory||0) + parseInt(subB?.practice||0);
        return totalB - totalA; 
    });

    for (const reg of sortedRegisters) {
        const sId = reg.subject_id || reg['subject id'];
        const gId = reg.group_id || reg['group id'];
        
        let subject = subjects.find(s => (s.subject_id || s['subject id']) === sId);
        if (!subject) {
            subject = { subject_name: '(ไม่พบข้อมูล)', theory: '0', practice: '2', credit: '0' };
        }

        const theory = parseInt(subject.theory || 0);
        const practice = parseInt(subject.practice || 0);
        const totalPeriodsNeeded = theory + practice;
        
        if (totalPeriodsNeeded === 0) continue;

        const isException = ['20001-1005', '30001-1003'].includes(sId);
        const isGenEdCode = (
            sId.startsWith('20000') || 
            sId.startsWith('20001') || 
            sId.startsWith('30000') || 
            sId.startsWith('30001')
        );
        const isGenEdTarget = isGenEdCode && !isException;
        const isPureTheory = (practice === 0);
        const isTargetTheory = isGenEdTarget || isPureTheory;

        let candidateRooms = [...rooms];
        if (isTargetTheory) {
            const theoryRooms = rooms.filter(r => {
                const rType = r.room_type ? r.room_type.trim() : '';
                return rType.toLowerCase() === 'theory';
            });
            if(theoryRooms.length > 0) candidateRooms = theoryRooms;
        }

        candidateRooms.sort((a, b) => {
            const idA = a.room_id || a['room id'];
            const idB = b.room_id || b['room id'];
            return (roomUsageCount[idA] || 0) - (roomUsageCount[idB] || 0);
        });

        let assignment = teachAssignments.find(a => (a.subject_id || a['subject id']) === sId && (a.group_id || a['group id']) === gId);
        if (!assignment) {
            assignment = teachAssignments.find(a => (a.subject_id || a['subject id']) === sId);
        }
        const teacherId = assignment ? (assignment.teacher_id || assignment['teacher id']) : null;

        let daysToCheck = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
        daysToCheck.sort((d1, d2) => {
            const load1 = getLoad(groupDailyLoad, gId, d1) + (teacherId ? getLoad(teacherDailyLoad, teacherId, d1) : 0);
            const load2 = getLoad(groupDailyLoad, gId, d2) + (teacherId ? getLoad(teacherDailyLoad, teacherId, d2) : 0);
            return load1 - load2;
        });

        let isBooked = false;
        const phases = [{ maxPeriod: 10 }, { maxPeriod: 12 }];

        for (const phase of phases) {
            if (isBooked) break;

            for (const day of daysToCheck) {
                if (isBooked) break;

                for (let startP = 1; startP <= (phase.maxPeriod - totalPeriodsNeeded + 1); startP++) {
                    let slotsFound = [];
                    let validBlock = true;

                    for (let k = 0; k < totalPeriodsNeeded; k++) {
                        const currentP = startP + k;
                        if (currentP === 5) { validBlock = false; break; }
                        if (phase.maxPeriod === 10 && currentP > 10) { validBlock = false; break; }

                        const slotObj = timeslots.find(ts => ts.day === day && parseInt(ts.period) === currentP);
                        if (!slotObj) { validBlock = false; break; }

                        if (!isSlotFree(slotObj.timeslot_id || slotObj['timeslot id'], teacherId, gId, null)) {
                            validBlock = false;
                            break;
                        }
                        slotsFound.push(slotObj);
                    }

                    if (validBlock) {
                        let assignedRoom = null;
                        for (const room of candidateRooms) {
                            const rId = room.room_id || room['room id'];
                            const roomIsFree = slotsFound.every(s => 
                                isSlotFree(s.timeslot_id || s['timeslot id'], null, null, rId)
                            );

                            if (roomIsFree) {
                                assignedRoom = rId;
                                break;
                            }
                        }

                        if (assignedRoom) {
                            slotsFound.forEach(slot => {
                                timetable.push({
                                    group_id: gId,
                                    timeslot_id: slot.timeslot_id || slot['timeslot id'],
                                    subject_id: sId,
                                    teacher_id: teacherId,
                                    room_id: assignedRoom
                                });
                            });

                            addLoad(groupDailyLoad, gId, day, totalPeriodsNeeded);
                            if (teacherId) addLoad(teacherDailyLoad, teacherId, day, totalPeriodsNeeded);
                            roomUsageCount[assignedRoom]++;

                            isBooked = true;
                            break;
                        }
                    }
                }
            }
        }

        if (!isBooked) {
            unassignedLog.push({
                subject_id: sId,
                subject_name: subject.subject_name,
                group_id: gId,
                teacher_id: teacherId,
                reason: 'No slot/room available'
            });
        }
    }
    return { timetable, unassignedLog };
}

// 3. ฟังก์ชันสร้าง HTML
function renderHTML(viewType, viewId, timetableData, timeslots, subjects, teachers, rooms, groups) {
    try {
        const { timetable } = timetableData;
        const daysEn = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
        const dayMap = { 'Mon': 'วันจันทร์', 'Tue': 'วันอังคาร', 'Wed': 'วันพุธ', 'Thu': 'วันพฤหัสบดี', 'Fri': 'วันศุกร์' };
        const periods = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
        
        const getPeriodTime = (p) => {
            const startHour = 8 + (p - 1);
            const endHour = 8 + p;
            const formatTime = (h) => (h < 10 ? '0' + h : h) + '.00';
            return `${formatTime(startHour)}-${formatTime(endHour)}`;
        };

        const groupOptions = groups.map(g => {
            const id = g.group_id || g['group id'];
            const name = g.group_name || g['group name'];
            return `<option value="${id}" ${viewType === 'group' && viewId === id ? 'selected' : ''}>${id} - ${name}</option>`;
        }).join('');

        const teacherOptions = teachers.map(t => {
            const id = t.teacher_id || t['teacher id'];
            const name = t.teacher_name || t['teacher name'];
            return `<option value="${id}" ${viewType === 'teacher' && viewId === id ? 'selected' : ''}>${name}</option>`;
        }).join('');

        const roomOptions = rooms.map(r => {
            const id = r.room_id || r['room id'];
            const name = r.room_name || r['room name'];
            return `<option value="${id}" ${viewType === 'room' && viewId === id ? 'selected' : ''}>${id} - ${name}</option>`;
        }).join('');

        const excelLink = `'/download-excel?type=${viewType || ''}&id=${viewId || ''}'`;

        let html = `
        <html>
        <head>
            <meta charset="UTF-8">
            <title>ระบบจัดตารางสอนอัตโนมัติ</title>
            <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;600&display=swap" rel="stylesheet">
            <style>
                body { font-family: 'Prompt', sans-serif; font-weight: 300; padding: 0; margin: 0; background-color: #f4f6f9; }
                .navbar { background-color: #004a99; padding: 15px 20px; color: white; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; box-shadow: 0 2px 5px rgba(0,0,0,0.2); }
                .navbar h1 { margin: 0; font-size: 20px; font-weight: 600; }
                .navbar h1 span { font-size: 12px; background: #FF5722; padding: 2px 8px; border-radius: 10px; margin-left: 10px; font-weight: 400; color: white;}
                .menu-container { display: flex; gap: 10px; align-items: center; }
                .menu-item { display: flex; flex-direction: column; }
                .menu-item label { font-size: 10px; margin-bottom: 2px; color: #bbdefb; }
                select { padding: 8px; border-radius: 4px; border: none; font-family: 'Prompt', sans-serif; font-size: 14px; min-width: 180px; }
                
                /* จัดปุ่มแนวนอน */
                .btn-group { 
                    display: flex; 
                    gap: 5px; 
                    margin-left: 10px; 
                    align-items: center; 
                }
                
                .btn { color: white; text-decoration: none; padding: 8px 12px; border-radius: 4px; font-size: 14px; cursor: pointer; border: none; font-family: 'Prompt'; display: flex; align-items: center; gap: 5px; }
                .btn:hover { opacity: 0.9; }
                .btn-excel { background: #2e7d32; }
                .btn-pdf { background: #d32f2f; }
                .btn-csv { background: #f57f17; }
                
                .print-select {
                    padding: 8px;
                    border-radius: 4px;
                    border: none;
                    font-family: 'Prompt', sans-serif;
                    font-size: 14px;
                    background-color: #006064; 
                    color: white; 
                    cursor: pointer;
                    min-width: 160px;
                }
                .print-select option { background-color: white; color: black; }
                
                .content { padding: 20px; max-width: 1280px; margin: 0 auto; background: white; min-height: 100vh; }
                
                table.timetable { width: 100%; border-collapse: collapse; margin-top: 10px; table-layout: fixed; }
                table.timetable th { background-color: #e3f2fd; color: #333; font-weight: 600; border: 1px solid #000; border-bottom: 2px solid #2196F3; padding: 5px; vertical-align: middle; height: 50px; }
                .time-display { display: block; font-size: 12px; font-weight: normal; margin-bottom: 2px; color: #004a99; }
                .period-display { display: block; font-size: 10px; font-weight: 300; color: #555; }
                table.timetable td { border: 1px solid #000; padding: 4px; text-align: center; font-size: 11px; height: 55px; vertical-align: middle; overflow: hidden; }
                
                .header-title-box { margin-top: 20px; margin-bottom: 10px; border-left: 5px solid #004a99; padding-left: 10px; display: flex; align-items: baseline; gap: 15px; }
                .header-title-box h2 { margin: 0; font-size: 24px; }
                .advisor-inline { font-size: 16px; color: #555; }
                .advisor-inline strong { color: #2e7d32; }

                .summary-container { display: flex; gap: 10px; margin-top: 15px; align-items: flex-start; }
                table.summary { width: 100%; border-collapse: collapse; }
                table.summary th, table.summary td { border: 1px solid #999; padding: 6px; text-align: center; font-size: 12px; height: 30px; white-space: nowrap; overflow: hidden; }
                table.summary th { background-color: #e0e0e0; font-weight: 600; }
                table.summary tfoot td { background-color: #f1f1f1; font-weight: 600; color: #333; border: 1px solid #999; height: 30px; }
                
                .cell-sub { font-weight: normal; font-size: 13px; display:block; }
                .cell-info { font-size: 11px; display:block; }
                .txt-teacher { color: #d32f2f; }
                .txt-room { color: #1976d2; }
                .txt-group { color: #388e3c; }
                
                .pdf-footer { display: none; margin-top: 20px; width: 100%; justify-content: space-around; font-size: 10px; align-items: flex-end; }
                .signature-box { text-align: center; width: 22%; display: flex; flex-direction: column; align-items: center; }
                .signature-line { margin-top: 20px; margin-bottom: 5px; border-bottom: 1px dotted #000; width: 100%; height: 1px; }
                .signature-role { margin-top: 5px; font-weight: 600; }

                @media print {
                    @page { size: A4 landscape; margin: 5mm; }
                    body { -webkit-print-color-adjust: exact; print-color-adjust: exact; background: white; }
                    .navbar, .btn-group, .menu-container { display: none !important; }
                    .content { padding: 0; margin: 0; max-width: 100%; width: 100%; box-shadow: none; }
                    
                    table.timetable, table.summary { border: 1px solid #000 !important; width: 100%; }
                    table.timetable th, table.timetable td, table.summary th, table.summary td { border: 1px solid #000 !important; }
                    table.timetable th { background-color: #e3f2fd !important; color: #000 !important; }
                    table.summary th { background-color: #ccc !important; color: #000 !important; }
                    
                    .pdf-footer { display: flex !important; page-break-inside: avoid; }
                    .header-title-box { border-left: none !important; padding-left: 0; margin-top: 0; }
                    .cell-sub { font-weight: normal !important; }
                    table.summary tr { height: 30px !important; }

                    .print-page { page-break-after: always; display: block; }
                    .print-page:last-child { page-break-after: auto; }
                }
            </style>
            <script>
                function navigate(type, id) { if(id) window.location.href = '/?type=' + type + '&id=' + id; }
                function exportPDF() { window.print(); }
                function downloadExcel() { window.location.href = ${excelLink}; }
            </script>
        </head>
        <body>
            <div class="navbar">
                <h1>📅 ระบบจัดตารางสอน <span>Beta</span></h1>
                <div class="menu-container">
                    <div class="menu-item"><label>กลุ่มเรียน</label><select onchange="navigate('group', this.value)"><option value="">-- เลือก --</option>${groupOptions}</select></div>
                    <div class="menu-item"><label>ครูผู้สอน</label><select onchange="navigate('teacher', this.value)"><option value="">-- เลือก --</option>${teacherOptions}</select></div>
                    <div class="menu-item"><label>ห้องเรียน</label><select onchange="navigate('room', this.value)"><option value="">-- เลือก --</option>${roomOptions}</select></div>
                    
                    <div class="btn-group">
                        <select class="print-select" onchange="if(this.value) window.location.href=this.value">
                            <option value="">🖨️ พิมพ์ทั้งหมด...</option>
                            <option value="/?type=group&id=all">พิมพ์ทุกกลุ่มเรียน</option>
                            <option value="/?type=teacher&id=all">พิมพ์ทุกครูผู้สอน</option>
                            <option value="/?type=room&id=all">พิมพ์ทุกห้องเรียน</option>
                        </select>
                        <button onclick="downloadExcel()" class="btn btn-excel">📗 Excel</button>
                        <button onclick="exportPDF()" class="btn btn-pdf">📕 PDF (Vector)</button>
                        <a href="/download-csv" class="btn btn-csv">📥 CSV</a>
                    </div>
                </div>
            </div>
            <div class="content" id="report-content">`;

        const renderSingleSchedule = (cType, cId) => {
            let output = "";
            let title = "", advisorSpan = "";
            
            if(cType === 'group') {
                const g = groups.find(x => (x.group_id || x['group id']) === cId);
                title = `ตารางเรียนกลุ่ม: ${g ? (g.group_name || g['group name']) : cId}`;
                if(g) {
                    const advId = g.advisor || g.advisor_id || g.teacher_id;
                    if(advId) {
                        const adv = teachers.find(t => (t.teacher_id || t['teacher id']) === advId);
                        advisorSpan = `<span class="advisor-inline"> ( <strong>ครูที่ปรึกษา:</strong> ${adv ? (adv.teacher_name || adv['teacher name']) : advId} )</span>`;
                    }
                }
            } else if(cType === 'teacher') {
                const t = teachers.find(x => (x.teacher_id || x['teacher id']) === cId);
                title = `ตารางสอน: ${t ? (t.teacher_name || t['teacher name']) : cId}`;
            } else {
                const r = rooms.find(x => (x.room_id || x['room id']) === cId);
                title = `ตารางการใช้ห้อง: ${r ? (r.room_name || r['room name']) : cId}`;
            }

            output += `<div class="header-title-box"><h2>${title}</h2>${advisorSpan}</div>`;
            
            output += `<table class="timetable"><thead><tr><th style="width:80px;">วัน/เวลา</th>${periods.map(p => `<th><span class="time-display">${getPeriodTime(p)}</span><span class="period-display">คาบที่ ${p}</span></th>`).join('')}</tr></thead><tbody>`;

            output += daysEn.map(day => {
                let rowHTML = `<tr><td class="day-col">${dayMap[day]}</td>`;
                for(let p=1; p<=12; p++) {
                    if(p===5) { rowHTML += `<td class="break-col">พัก</td>`; continue; }
                    
                    const getEntry = (pN) => {
                        const slot = timeslots.find(ts => ts.day === day && parseInt(ts.period) === pN);
                        if(!slot) return null;
                        return timetable.find(t => {
                            if(t.timeslot_id !== slot.timeslot_id && t.timeslot_id !== slot['timeslot id']) return false;
                            if(cType==='group' && t.group_id!==cId) return false;
                            if(cType==='teacher' && t.teacher_id!==cId) return false;
                            if(cType==='room' && t.room_id!==cId) return false;
                            return true;
                        });
                    };

                    const entry = getEntry(p);
                    if(entry) {
                        let span = 1;
                        for(let nextP=p+1; nextP<=12; nextP++) {
                            if(nextP===5) break;
                            const nextEntry = getEntry(nextP);
                            if(nextEntry && nextEntry.subject_id === entry.subject_id && nextEntry.room_id === entry.room_id && nextEntry.teacher_id === entry.teacher_id && nextEntry.group_id === entry.group_id) {
                                span++;
                            } else { break; }
                        }

                        const t = teachers.find(x => (x.teacher_id || x['teacher id']) === entry.teacher_id);
                        const r = rooms.find(x => (x.room_id || x['room id']) === entry.room_id);
                        let tName = t ? (t.teacher_name || t['teacher name']) : entry.teacher_id;
                        if(tName && tName.startsWith('ครู')) tName = tName.replace('ครู','');
                        const rName = r ? (r.room_name || r['room name']) : entry.room_id;

                        let info = '';
                        if(cType==='group') info = `<span class="cell-info txt-teacher">${tName}</span><span class="cell-info txt-room">ห้อง ${rName}</span>`;
                        else if(cType==='teacher') info = `<span class="cell-info txt-group">กลุ่ม ${entry.group_id}</span><span class="cell-info txt-room">ห้อง ${rName}</span>`;
                        else info = `<span class="cell-info txt-teacher">${tName}</span><span class="cell-info txt-group">กลุ่ม ${entry.group_id}</span>`;

                        rowHTML += `<td colspan="${span}" style="background:#e3f2fd;"><div class="cell-content"><span class="cell-sub">${entry.subject_id}</span>${info}</div></td>`;
                        p += (span-1);
                    } else {
                        rowHTML += `<td></td>`;
                    }
                }
                return rowHTML + `</tr>`;
            }).join('');
            output += `</tbody></table>`;

            const relevantEntries = timetable.filter(t => {
                if(cType==='group') return t.group_id===cId;
                if(cType==='teacher') return t.teacher_id===cId;
                if(cType==='room') return t.room_id===cId;
                return false;
            });
            const grouped = {};
            relevantEntries.forEach(e => {
                if(!grouped[e.subject_id]) grouped[e.subject_id] = { id: e.subject_id, count: 0 };
                grouped[e.subject_id].count++;
            });
            
            let allRows = Object.values(grouped).map((item, idx) => {
                let sub = subjects.find(s => (s.subject_id || s['subject id']) === item.id);
                if(!sub) sub = {subject_name:'(ไม่พบข้อมูล)', theory:0, practice:0, credit:0};
                return { index: idx+1, id: item.id, name: sub.subject_name, t: parseInt(sub.theory||0), p: parseInt(sub.practice||0), c: parseInt(sub.credit||0), total: item.count };
            });

            const grandTotal = allRows.reduce((acc, curr) => ({
                t: acc.t + curr.t, p: acc.p + curr.p, c: acc.c + curr.c, total: acc.total + curr.total
            }), { t:0, p:0, c:0, total:0 });

            const grandTotalRow = `<tr style="background-color:#e0e0e0; font-weight:bold;"><td colspan="3" style="text-align:right;">รวมทั้งสิ้น</td><td>${grandTotal.t}</td><td>${grandTotal.p}</td><td>${grandTotal.c}</td><td>${grandTotal.total}</td></tr>`;

            let leftRows, rightRows;
            let leftFooterHTML, rightFooterHTML;

            if (allRows.length <= 3) {
                leftRows = allRows;
                rightRows = [];
                const leftSum = leftRows.reduce((a,b)=>({t:a.t+b.t, p:a.p+b.p, c:a.c+b.c, total:a.total+b.total}), {t:0,p:0,c:0,total:0});
                leftFooterHTML = `<tfoot>
                    <tr><td colspan="3" style="text-align:right;">รวมย่อย</td><td>${leftSum.t}</td><td>${leftSum.p}</td><td>${leftSum.c}</td><td>${leftSum.total}</td></tr>
                    ${grandTotalRow}
                </tfoot>`;
                rightFooterHTML = '';
            } else {
                const mid = Math.ceil(allRows.length/2);
                leftRows = allRows.slice(0, mid);
                rightRows = allRows.slice(mid);

                const leftSum = leftRows.reduce((a,b)=>({t:a.t+b.t, p:a.p+b.p, c:a.c+b.c, total:a.total+b.total}), {t:0,p:0,c:0,total:0});
                const rightSum = rightRows.reduce((a,b)=>({t:a.t+b.t, p:a.p+b.p, c:a.c+b.c, total:a.total+b.total}), {t:0,p:0,c:0,total:0});

                leftFooterHTML = `<tfoot><tr><td colspan="3" style="text-align:right;">รวมย่อย</td><td>${leftSum.t}</td><td>${leftSum.p}</td><td>${leftSum.c}</td><td>${leftSum.total}</td></tr></tfoot>`;
                rightFooterHTML = `<tfoot><tr><td colspan="3" style="text-align:right;">รวมย่อย</td><td>${rightSum.t}</td><td>${rightSum.p}</td><td>${rightSum.c}</td><td>${rightSum.total}</td></tr>${grandTotalRow}</tfoot>`;
            }
            
            const renderRows = (rows) => rows.map(r => `<tr><td>${r.index}</td><td style="text-align:left;">${r.id}</td><td style="text-align:left;">${r.name}</td><td>${r.t}</td><td>${r.p}</td><td>${r.c}</td><td>${r.total}</td></tr>`).join('');
            
            output += `<div class="summary-container">
                <table class="summary" style="width:50%;">
                    <thead><tr><th>ลำดับ</th><th>รหัสวิชา</th><th>ชื่อวิชา</th><th>ท.</th><th>ป.</th><th>นก.</th><th>รวม</th></tr></thead>
                    <tbody>${renderRows(leftRows)}</tbody>
                    ${leftFooterHTML}
                </table>`;
            
            if (rightRows.length > 0) {
                output += `<table class="summary" style="width:50%;">
                    <thead><tr><th>ลำดับ</th><th>รหัสวิชา</th><th>ชื่อวิชา</th><th>ท.</th><th>ป.</th><th>นก.</th><th>รวม</th></tr></thead>
                    <tbody>${renderRows(rightRows)}</tbody>
                    ${rightFooterHTML}
                </table>`;
            } else {
                output += `<div style="width:50%;"></div>`;
            }
            output += `</div>`;

            output += `<div class="pdf-footer">
                <div class="signature-box"><div>ตรวจแล้ว</div><div class="signature-line"></div><div class="signature-role">หัวหน้าแผนกวิชา</div></div>
                <div class="signature-box"><div>ตรวจแล้ว</div><div class="signature-line"></div><div class="signature-role">หัวหน้างานพัฒนาหลักสูตรฯ</div></div>
                <div class="signature-box"><div>เห็นควรอนุญาต</div><div class="signature-line"></div><div class="signature-role">รองผู้อำนวยการฝ่ายวิชาการ</div></div>
                <div class="signature-box"><div>อนุญาต</div><div class="signature-line"></div><div class="signature-role">ผู้อำนวยการวิทยาลัย...</div></div>
            </div>`;

            return output;
        };

        if (!viewType) {
            html += `<div style="text-align:center; margin-top:50px; color:#666;"><h2>ยินดีต้อนรับ</h2><p>กรุณาเลือกข้อมูลจากเมนูด้านบน</p></div>`;
        } else if (viewId === 'all') {
            let items = [];
            if(viewType === 'group') items = groups.map(g => g.group_id || g['group id']);
            else if(viewType === 'teacher') items = teachers.map(t => t.teacher_id || t['teacher id']);
            else if(viewType === 'room') items = rooms.map(r => r.room_id || r['room id']);

            items.forEach(itemId => {
                html += `<div class="print-page">`;
                html += renderSingleSchedule(viewType, itemId);
                html += `</div>`;
            });

        } else {
            html += renderSingleSchedule(viewType, viewId);
        }

        html += `</div></body></html>`;
        return html;
    } catch(e) {
        return `<h1>Error Rendering HTML</h1><pre>${e.message}</pre>`;
    }
}

app.get('/', async (req, res) => {
    try {
        const [teachers, rooms, subjects, timeslots, registers, teachAssignments, groups] = await Promise.all([
            readCSV('teacher.csv'), readCSV('room.csv'), readCSV('subject.csv'),
            readCSV('timeslot.csv'), readCSV('register.csv'), readCSV('teach.csv'), readCSV('student_group.csv')
        ]);
        const timetableRes = solveTimetable(teachers, rooms, subjects, timeslots, registers, teachAssignments, groups);
        
        const header = 'group_id,timeslot_id,subject_id,teacher_id,room_id\n';
        const rows = timetableRes.timetable.map(d => `${d.group_id},${d.timeslot_id},${d.subject_id},${d.teacher_id},${d.room_id}`).join('\n');
        fs.writeFileSync('output.csv', '\ufeff' + header + rows, 'utf8');

        const { type, id } = req.query;
        res.send(renderHTML(type, id, timetableRes, timeslots, subjects, teachers, rooms, groups));
    } catch (e) {
        console.error(e);
        res.status(500).send("Server Error: " + e.message);
    }
});

app.get('/download-csv', (req, res) => {
    const file = path.join(__dirname, 'output.csv');
    if(fs.existsSync(file)) res.download(file); else res.status(404).send("File not found");
});

app.get('/download-excel', async (req, res) => {
    try {
        const { type, id } = req.query;
        if(!type || !id || id === 'all') return res.send('Excel export supports single schedule only for now.');

        const [teachers, rooms, subjects, timeslots, registers, teachAssignments, groups] = await Promise.all([
            readCSV('teacher.csv'), readCSV('room.csv'), readCSV('subject.csv'),
            readCSV('timeslot.csv'), readCSV('register.csv'), readCSV('teach.csv'), readCSV('student_group.csv')
        ]);
        const { timetable } = solveTimetable(teachers, rooms, subjects, timeslots, registers, teachAssignments, groups);

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Timetable');

        sheet.pageSetup = { paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 1 };

        let title = "", advisorName = "";
        if(type === 'group') {
            const g = groups.find(x => (x.group_id || x['group id']) === id);
            title = `ตารางเรียนกลุ่ม: ${g ? (g.group_name || g['group name']) : id}`;
            const advId = g ? (g.advisor || g.advisor_id || g.teacher_id) : "";
            if(advId) {
                const adv = teachers.find(t => (t.teacher_id || t['teacher id']) === advId);
                advisorName = adv ? (adv.teacher_name || adv['teacher name']) : advId;
            }
        } else if(type === 'teacher') {
            const t = teachers.find(x => (x.teacher_id || x['teacher id']) === id);
            title = `ตารางสอน: ${t ? (t.teacher_name || t['teacher name']) : id}`;
        } else {
            const r = rooms.find(x => (x.room_id || x['room id']) === id);
            title = `ตารางการใช้ห้อง: ${r ? (r.room_name || r['room name']) : id}`;
        }

        sheet.mergeCells('A1:M1');
        const titleCell = sheet.getCell('A1');
        titleCell.value = title;
        titleCell.font = { name: 'Prompt', size: 16, bold: true };
        titleCell.alignment = { horizontal: 'center' };

        if(advisorName) {
            sheet.mergeCells('A2:M2');
            const advCell = sheet.getCell('A2');
            advCell.value = `ครูที่ปรึกษา: ${advisorName}`;
            advCell.font = { name: 'Prompt', size: 12 };
            advCell.alignment = { horizontal: 'center' };
        }

        const startRow = advisorName ? 4 : 3;
        
        const getPeriodTime = (p) => {
            const startHour = 8 + (p - 1);
            const endHour = 8 + p;
            const formatTime = (h) => (h < 10 ? '0' + h : h) + '.00';
            return `${formatTime(startHour)}-${formatTime(endHour)}`;
        };
        
        const headerValues = ['วัน/คาบ', ...Array.from({length:12}, (_,i) => {
            return `${getPeriodTime(i+1)}\n(คาบ ${i+1})`; 
        })];
        
        const headerRow = sheet.getRow(startRow);
        headerRow.values = headerValues;
        headerRow.height = 40; 
        
        headerRow.font = { name: 'Prompt', bold: false, size: 10 };
        headerRow.eachCell(cell => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E3F2FD' } };
            cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
            cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        });

        const daysEn = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
        const dayMap = { 'Mon': 'วันจันทร์', 'Tue': 'วันอังคาร', 'Wed': 'วันพุธ', 'Thu': 'วันพฤหัสบดี', 'Fri': 'วันศุกร์' };
        let currentRow = startRow + 1;

        daysEn.forEach(day => {
            const row = sheet.getRow(currentRow);
            row.getCell(1).value = dayMap[day];
            row.getCell(1).font = { name: 'Prompt', bold: true };
            row.getCell(1).border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
            row.getCell(1).alignment = { horizontal: 'center', vertical: 'middle' };

            for(let p=1; p<=12; p++) {
                const cell = row.getCell(p+1);
                cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.font = { name: 'Prompt', size: 10, bold: false };

                if(p === 5) {
                    cell.value = 'พัก';
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'EEEEEE' } };
                    cell.alignment = { textRotation: 90, vertical: 'middle', horizontal: 'center' };
                    continue;
                }

                const slot = timeslots.find(ts => ts.day === day && parseInt(ts.period) === p);
                if(slot) {
                    const entry = timetable.find(t => {
                        if(t.timeslot_id !== slot.timeslot_id && t.timeslot_id !== slot['timeslot id']) return false;
                        if(type === 'group' && t.group_id !== id) return false;
                        if(type === 'teacher' && t.teacher_id !== id) return false;
                        if(type === 'room' && t.room_id !== id) return false;
                        return true;
                    });

                    if(entry) {
                        let span = 1;
                        for(let nextP=p+1; nextP<=12; nextP++) {
                            if(nextP===5) break;
                            const nextSlot = timeslots.find(ts => ts.day === day && parseInt(ts.period) === nextP);
                            if(!nextSlot) break;
                            const nextEntry = timetable.find(t => {
                                if(t.timeslot_id !== nextSlot.timeslot_id && t.timeslot_id !== nextSlot['timeslot id']) return false;
                                if(type === 'group' && t.group_id !== id) return false;
                                if(type === 'teacher' && t.teacher_id !== id) return false;
                                if(type === 'room' && t.room_id !== id) return false;
                                return true;
                            });
                            if(nextEntry && nextEntry.subject_id === entry.subject_id) span++; else break;
                        }

                        const sub = subjects.find(s => (s.subject_id || s['subject id']) === entry.subject_id);
                        const tea = teachers.find(t => (t.teacher_id || t['teacher id']) === entry.teacher_id);
                        const rm = rooms.find(r => (r.room_id || r['room id']) === entry.room_id);
                        
                        let cellText = `${entry.subject_id}\n`;
                        let teaName = tea ? (tea.teacher_name || tea['teacher name']) : entry.teacher_id;
                        if(teaName && teaName.startsWith('ครู')) teaName = teaName.replace('ครู','');

                        if(type === 'group') cellText += `${teaName}\nห้อง ${rm ? (rm.room_name || rm['room name']) : entry.room_id}`;
                        else if(type === 'teacher') cellText += `กลุ่ม ${entry.group_id}\nห้อง ${rm ? (rm.room_name || rm['room name']) : entry.room_id}`;
                        else cellText += `${teaName}\nกลุ่ม ${entry.group_id}`;
                        
                        cell.value = cellText;
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E3F2FD' } };
                        cell.font = { name: 'Prompt', size: 10, bold: false };

                        if(span > 1) {
                            sheet.mergeCells(currentRow, p+1, currentRow, p+span);
                        }
                        p += (span - 1);
                    }
                }
            }
            row.height = 50;
            currentRow++;
        });

        currentRow += 2;
        const sumHeader = sheet.getRow(currentRow);
        sumHeader.values = ['ลำดับ', 'รหัสวิชา', 'ชื่อวิชา', 'ท.', 'ป.', 'นก.', 'รวม'];
        sumHeader.font = { name: 'Prompt', bold: true };
        for(let i=1; i<=7; i++) {
            sumHeader.getCell(i).border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
            sumHeader.getCell(i).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E0E0E0' } };
            sumHeader.getCell(i).alignment = { horizontal: 'center' };
        }
        currentRow++;

        const relevantEntries = timetable.filter(t => {
            if(type === 'group') return t.group_id === id;
            if(type === 'teacher') return t.teacher_id === id;
            if(type === 'room') return t.room_id === id;
            return false;
        });
        const groupedData = {};
        relevantEntries.forEach(e => {
            if(!groupedData[e.subject_id]) groupedData[e.subject_id] = { id: e.subject_id, count: 0 };
            groupedData[e.subject_id].count++;
        });

        let idx = 1, totalT=0, totalP=0, totalC=0, totalAll=0;
        let allRows = Object.values(groupedData).map((item, index) => {
            let sub = subjects.find(s => (s.subject_id || s['subject id']) === item.id);
            if(!sub) sub = {subject_name:'(ไม่พบข้อมูล)', theory:0, practice:0, credit:0};
            return { index: index+1, id: item.id, name: sub.subject_name, t: parseInt(sub.theory||0), p: parseInt(sub.practice||0), c: parseInt(sub.credit||0), total: item.count };
        });

        let leftRows, rightRows;
        if (allRows.length <= 3) {
            leftRows = allRows;
            rightRows = [];
        } else {
            const mid = Math.ceil(allRows.length / 2);
            leftRows = allRows.slice(0, mid);
            rightRows = allRows.slice(mid);
        }

        const addSummaryRows = (rows, startCol) => {
            let localRow = currentRow;
            rows.forEach(item => {
                const r = sheet.getRow(localRow);
                r.getCell(startCol).value = item.index;
                r.getCell(startCol+1).value = item.id;
                r.getCell(startCol+2).value = item.name;
                r.getCell(startCol+3).value = item.t;
                r.getCell(startCol+4).value = item.p;
                r.getCell(startCol+5).value = item.c;
                r.getCell(startCol+6).value = item.total;

                for(let i=0; i<=6; i++) {
                    const c = r.getCell(startCol+i);
                    c.font = { name: 'Prompt', size: 10 };
                    c.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
                    c.alignment = { horizontal: 'center' };
                    if(i === 2) c.alignment = { horizontal: 'left' };
                }
                
                totalT+=item.t; totalP+=item.p; totalC+=item.c; totalAll+=item.total;
                localRow++;
            });
            return localRow;
        };

        const leftEndRow = addSummaryRows(leftRows, 1);
        
        if(rightRows.length > 0) {
            const sumHeaderRight = sheet.getRow(currentRow-1);
            ['ลำดับ', 'รหัสวิชา', 'ชื่อวิชา', 'ท.', 'ป.', 'นก.', 'รวม'].forEach((val, i) => {
                const c = sumHeaderRight.getCell(8+i);
                c.value = val;
                c.font = { name: 'Prompt', bold: true };
                c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E0E0E0' } };
                c.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
                c.alignment = { horizontal: 'center' };
            });
            addSummaryRows(rightRows, 8);
        }

        currentRow = Math.max(leftEndRow, currentRow + rightRows.length);

        const footerRow = sheet.getRow(currentRow);
        sheet.mergeCells(currentRow, 1, currentRow, 3);
        footerRow.getCell(1).value = 'รวมทั้งสิ้น';
        footerRow.getCell(4).value = totalT;
        footerRow.getCell(5).value = totalP;
        footerRow.getCell(6).value = totalC;
        footerRow.getCell(7).value = totalAll;
        
        footerRow.font = { name: 'Prompt', bold: true };
        [1,4,5,6,7].forEach(c => {
            const cell = footerRow.getCell(c);
            cell.border = { top:{style:'thin'}, left:{style:'thin'}, bottom:{style:'thin'}, right:{style:'thin'} };
            cell.alignment = { horizontal: 'center' };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E0E0E0' } };
        });

        currentRow += 4;
        const sigRow1 = sheet.getRow(currentRow);
        sigRow1.values = ['', 'ตรวจแล้ว', '', 'ตรวจแล้ว', '', 'เห็นควรอนุญาต', '', '', 'อนุญาต'];
        const sigRowLine = sheet.getRow(currentRow + 2);
        sigRowLine.values = ['', '................................', '', '................................', '', '................................', '', '', '................................'];
        const sigRowRole = sheet.getRow(currentRow + 3);
        sigRowRole.values = ['', 'หัวหน้าแผนกวิชา', '', 'หัวหน้างานพัฒนาหลักสูตรฯ', '', 'รองผู้อำนวยการฝ่ายวิชาการ', '', '', 'ผู้อำนวยการวิทยาลัย...'];

        [sigRow1, sigRowLine, sigRowRole].forEach(r => {
            r.font = { name: 'Prompt', size: 10 };
            r.alignment = { horizontal: 'center' };
        });

        sheet.getColumn(2).width = 25;
        sheet.getColumn(4).width = 25;
        sheet.getColumn(6).width = 25;
        sheet.getColumn(9).width = 25;

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=timetable.xlsx');
        await workbook.xlsx.write(res);
        res.end();

    } catch (e) {
        res.status(500).send("Excel Error: " + e.message);
    }
});

app.listen(PORT, () => {
    console.log(`\n[SUCCESS] Server Running at http://localhost:${PORT}`);
});