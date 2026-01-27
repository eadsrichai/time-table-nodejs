const express = require('express');
const fs = require('fs');
const csv = require('csv-parser');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 80;

// 1. ฟังก์ชันโหลดข้อมูล (ตัดช่องว่าง Key/Value อัตโนมัติ)
async function readCSV(filePath) {
    const results = [];
    return new Promise((resolve) => {
        if (!fs.existsSync(filePath)) return resolve([]);
        fs.createReadStream(filePath)
            .pipe(csv())
            .on('data', (data) => {
                const cleanData = {};
                Object.keys(data).forEach(key => { 
                    cleanData[key.trim()] = data[key] ? data[key].trim() : ''; 
                });
                results.push(cleanData);
            })
            .on('end', () => resolve(results))
            .on('error', (err) => resolve([]));
    });
}

// 2. Logic จัดตาราง (No Constraints / Free Flow ตามไฟล์ Constra_Basic)
function solveTimetable(teachers, rooms, subjects, timeslots, registers, teachAssignments, groups) {
    let timetable = [];
    let unassignedLog = [];

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

    const validRegisters = registers.filter(r => {
        const sId = r.subject_id || r['subject id'];
        return subjects.some(s => (s.subject_id || s['subject id']) === sId);
    });

    for (const reg of validRegisters) {
        const sId = reg.subject_id || reg['subject id'];
        const gId = reg.group_id || reg['group id'];
        const subject = subjects.find(s => (s.subject_id || s['subject id']) === sId);
        
        if (!subject) continue;

        const totalPeriodsNeeded = parseInt(subject.theory || 0) + parseInt(subject.practice || 0);
        
        let assignment = teachAssignments.find(a => (a.subject_id || a['subject id']) === sId && (a.group_id || a['group id']) === gId);
        if (!assignment) {
            assignment = teachAssignments.find(a => (a.subject_id || a['subject id']) === sId);
        }
        const teacherId = assignment ? (assignment.teacher_id || assignment['teacher id']) : null;

        let periodsAssigned = 0;

        // First Fit Strategy (ใส่ทันทีที่ว่าง)
        for (let i = 0; i < totalPeriodsNeeded; i++) {
            for (const slot of timeslots) {
                const tSlotId = slot.timeslot_id || slot['timeslot id'];
                const period = parseInt(slot.period);

                if (period === 5) continue; // เว้นพักเที่ยง
                if (!isSlotFree(tSlotId, teacherId, gId, null)) continue;

                let assignedRoom = null;
                for (const room of rooms) {
                    const rId = room.room_id || room['room id'];
                    if (isSlotFree(tSlotId, null, null, rId)) {
                        assignedRoom = rId;
                        break;
                    }
                }

                if (assignedRoom) {
                    timetable.push({
                        group_id: gId,
                        timeslot_id: tSlotId,
                        subject_id: sId,
                        teacher_id: teacherId,
                        room_id: assignedRoom
                    });
                    periodsAssigned++;
                    break;
                }
            }
        }

        if (periodsAssigned < totalPeriodsNeeded) {
            unassignedLog.push({
                subject_id: sId,
                subject_name: subject.subject_name,
                group_id: gId,
                teacher_id: teacherId,
                missing: totalPeriodsNeeded - periodsAssigned,
                reason: 'No available slot/room'
            });
        }
    }
    return { timetable, unassignedLog };
}

// 3. ฟังก์ชันสร้าง HTML
function renderHTML(viewType, viewId, timetableData, timeslots, subjects, teachers, rooms, groups) {
    const { timetable, unassignedLog } = timetableData;
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
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
        
        <style>
            body { font-family: 'Prompt', sans-serif; font-weight: 300; padding: 0; margin: 0; background-color: #f4f6f9; }
            .navbar { background-color: #004a99; padding: 15px 20px; color: white; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; box-shadow: 0 2px 5px rgba(0,0,0,0.2); }
            .navbar h1 { margin: 0; font-size: 20px; font-weight: 600; display: flex; align-items: center; }
            .navbar h1 span { font-size: 12px; background: #FF5722; padding: 2px 8px; border-radius: 10px; margin-left: 10px; font-weight: 400; color: white;}
            .menu-container { display: flex; gap: 10px; align-items: center; }
            .menu-item { display: flex; flex-direction: column; }
            .menu-item label { font-size: 10px; margin-bottom: 2px; color: #bbdefb; }
            select { padding: 8px; border-radius: 4px; border: none; font-family: 'Prompt', sans-serif; font-weight: 300; font-size: 14px; min-width: 180px; cursor: pointer; }
            select:focus { outline: 2px solid #82b1ff; }
            .btn-group { display: flex; gap: 5px; margin-left: 10px; }
            .btn { color: white; text-decoration: none; padding: 8px 12px; border-radius: 4px; font-size: 14px; transition: 0.3s; cursor: pointer; border: none; display: flex; align-items: center; gap: 5px; font-family: 'Prompt', sans-serif; font-weight: 300; }
            .btn-excel { background: #2e7d32; }
            .btn-pdf { background: #d32f2f; }
            .btn-csv { background: #f57f17; }
            .content { padding: 20px; max-width: 1280px; margin: 0 auto; padding-bottom: 50px; background: white; }
            
            table.timetable { width: 100%; border-collapse: collapse; margin-top: 10px; table-layout: fixed; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
            table.timetable th, table.timetable td { border: 1px solid #000; padding: 4px; text-align: center; font-size: 11px; height: 55px; vertical-align: middle; overflow: hidden; }
            table.timetable th { background-color: #e3f2fd; color: #333; font-weight: 600; border-bottom: 2px solid #2196F3; }
            
            /* Header & Advisor Styles */
            .header-title-box {
                margin-top: 20px;
                margin-bottom: 10px;
                border-left: 5px solid #004a99;
                padding-left: 10px;
                display: flex;
                align-items: baseline; 
                flex-wrap: wrap;
                gap: 15px;
            }
            .header-title-box h2 {
                margin: 0;
                color: #333;
                font-size: 24px;
                font-weight: 600;
            }
            .advisor-inline {
                font-size: 16px;
                color: #555;
            }
            .advisor-inline strong {
                color: #2e7d32;
                font-weight: 600;
            }

            .summary-container { display: flex; gap: 10px; margin-top: 15px; align-items: flex-start; }
            table.summary { width: 100%; border-collapse: collapse; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
            table.summary th, table.summary td { border: 1px solid #999; padding: 6px; text-align: center; font-size: 12px; }
            table.summary th { background-color: #e0e0e0; color: #333; height: 35px; font-weight: 600; border-bottom: 2px solid #999; }
            table.summary tfoot td { background-color: #f1f1f1; font-weight: 600; color: #333; }
            
            .grand-total-box { background-color: #f5f5f5; border: 2px solid #ccc; border-radius: 5px; padding: 10px; text-align: center; margin-top: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); display: flex; justify-content: space-around; page-break-inside: avoid; }
            .grand-total-box .total-item { color: #333; font-weight: 400; }
            .grand-total-box .total-item b { color: #000; font-weight: 600; }

            .group-legend { margin-top: 20px; border-top: 1px dashed #ccc; padding-top: 10px; font-size: 12px; color: #555; }
            .group-legend h4 { margin: 0 0 5px 0; color: #333; font-weight: 600; }
            .group-legend ul { list-style: none; padding: 0; margin: 0; display: flex; flex-wrap: wrap; gap: 15px; }
            .group-legend li { background: #f5f5f5; padding: 2px 8px; border-radius: 4px; border: 1px solid #eee; }
            .group-legend strong { color: #388e3c; font-weight: 600; }

            .error-log { margin-top: 30px; background: #ffebee; border: 1px solid #ef5350; padding: 15px; border-radius: 5px; color: #c62828; }
            
            .time-header { font-size: 12px; display: block; margin-bottom: 2px; font-weight: 600; }
            .period-label { font-size: 9px; color: #666; font-weight: 300; }
            .day-col { background-color: #fafafa; font-weight: 600; width: 80px; color: #004a99; border-right: 2px solid #ddd; }
            .break-col { background-color: #eee; font-weight: 600; color: #777; vertical-align: middle; text-align: center; white-space: nowrap; }
            .cell-content { display: flex; flex-direction: column; justify-content: center; height: 100%; }
            .cell-sub { font-weight: 600; font-size: 13px; color: #000; margin-bottom: 3px; }
            .cell-info { font-size: 11px; margin-top: 2px; }
            .txt-teacher { color: #d32f2f; }
            .txt-room { color: #1976d2; }
            .txt-group { color: #388e3c; }
            .welcome-box { text-align: center; margin-top: 50px; color: #666; }

            /* --- Footer Signature Block (เฉพาะใน PDF) --- */
            .pdf-footer {
                display: none; /* ซ่อนบนเว็บ */
                margin-top: 20px;
                width: 100%;
                justify-content: space-around;
                align-items: flex-end;
                font-size: 10px;
                color: #000;
            }
            .signature-box {
                text-align: center;
                width: 22%; /* ปรับให้พอดี 4 ช่อง */
                display: flex;
                flex-direction: column;
                align-items: center;
            }
            .signature-line {
                margin-top: 20px;
                margin-bottom: 5px;
                border-bottom: 1px dotted #000;
                width: 100%;
                height: 1px;
            }
            .signature-role {
                margin-top: 5px;
                font-weight: 600;
            }

            /* --- Print / PDF Styles --- */
            .fit-to-page { width: 100% !important; padding: 0 !important; }
            
            /* แสดง Footer เมื่อเป็น PDF */
            .fit-to-page .pdf-footer { display: flex !important; }

            .fit-to-page .header-title-box { margin-top: 5px !important; margin-bottom: 2px !important; padding-left: 5px !important; border-left-width: 3px !important; }
            .fit-to-page h2 { font-size: 16px !important; }
            .fit-to-page .advisor-inline { font-size: 12px !important; }
            
            .fit-to-page table.timetable th, .fit-to-page table.timetable td { font-size: 9px !important; padding: 2px !important; height: 35px !important; }
            .fit-to-page .cell-sub { font-size: 10px !important; margin-bottom: 1px !important; }
            .fit-to-page .cell-info { font-size: 8px !important; }
            .fit-to-page .time-header { font-size: 9px !important; }
            .fit-to-page .period-label { display: none !important; } 
            .fit-to-page .summary-container { margin-top: 5px !important; gap: 5px !important; }
            .fit-to-page table.summary th, .fit-to-page table.summary td { font-size: 9px !important; padding: 2px !important; height: 20px !important; }
            .fit-to-page .grand-total-box { margin-top: 5px !important; padding: 5px !important; }
            .fit-to-page .grand-total-box .total-item { font-size: 10px !important; }
            .fit-to-page .group-legend { margin-top: 5px !important; padding-top: 5px !important; font-size: 9px !important; }
            .fit-to-page .group-legend h4 { margin-bottom: 2px !important; font-size: 10px !important;}
            .fit-to-page .group-legend li { padding: 1px 4px !important; }
            .fit-to-page .error-log { display: none !important; }

            @media print { .navbar { display: none; } .content { padding: 0; max-width: 100%; } table { page-break-inside: auto; } }
        </style>
        <script>
            function navigate(type, id) { if(id) window.location.href = '/?type=' + type + '&id=' + id; }
            function exportPDF() {
                const btn = document.querySelector('.btn-pdf');
                const element = document.getElementById('report-content');
                const originalText = btn.innerText;
                element.classList.add('fit-to-page');
                btn.innerText = 'กำลังสร้าง...';
                const opt = { margin: 5, filename: 'timetable-report.pdf', image: { type: 'jpeg', quality: 0.98 }, html2canvas: { scale: 2, useCORS: true }, jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' } };
                html2pdf().set(opt).from(element).output('bloburl').then(function(pdfUrl) {
                    window.open(pdfUrl, '_blank');
                    element.classList.remove('fit-to-page');
                    btn.innerText = originalText;
                }).catch(err => {
                    console.error(err);
                    alert('Error generating PDF');
                    element.classList.remove('fit-to-page');
                    btn.innerText = originalText;
                });
            }
            function downloadExcel() { window.location.href = ${excelLink}; }
        </script>
    </head>
    <body>
        <div class="navbar" id="navbar">
            <h1>📅 ระบบจัดตารางสอน <span>Beta</span></h1>
            <div class="menu-container">
                <div class="menu-item">
                    <label>เลือกกลุ่มเรียน</label>
                    <select onchange="navigate('group', this.value)">
                        <option value="">-- เลือกกลุ่มเรียน --</option>
                        ${groupOptions}
                    </select>
                </div>
                <div class="menu-item">
                    <label>เลือกครูผู้สอน</label>
                    <select onchange="navigate('teacher', this.value)">
                        <option value="">-- เลือกครูผู้สอน --</option>
                        ${teacherOptions}
                    </select>
                </div>
                <div class="menu-item">
                    <label>เลือกห้องเรียน</label>
                    <select onchange="navigate('room', this.value)">
                        <option value="">-- เลือกห้องเรียน --</option>
                        ${roomOptions}
                    </select>
                </div>
                <div class="btn-group">
                    <button onclick="downloadExcel()" class="btn btn-excel">📗 Excel</button>
                    <button onclick="exportPDF()" class="btn btn-pdf">📕 PDF (Preview)</button>
                    <a href="/download-csv" class="btn btn-csv">📥 CSV</a>
                </div>
            </div>
        </div>

        <div class="content" id="report-content">`;

    if (!viewType) {
        html += `
            <div class="welcome-box">
                <h2 style="color:#004a99; font-size:28px;">ยินดีต้อนรับ</h2>
                <p>กรุณาเลือก <b>กลุ่มเรียน, ครูผู้สอน หรือ ห้องเรียน</b> จากเมนูด้านบนเพื่อดูตาราง</p>
                <div style="font-size:60px; margin-top:20px;">👆</div>
            </div>`;
    } else {
        let title = "";
        let advisorSpan = "";

        if(viewType === 'group') {
            const currentGroup = groups.find(g => (g.group_id || g['group id']) === viewId);
            const gName = currentGroup ? (currentGroup.group_name || currentGroup['group name']) : viewId;
            title = `ตารางเรียนกลุ่ม: ${gName}`;
            
            if(currentGroup) {
                const advisorId = currentGroup.advisor || currentGroup['advisor'] || 
                                   currentGroup.advisor_id || currentGroup['advisor id'] || 
                                   currentGroup.teacher_id || currentGroup['teacher id'];
                
                if(advisorId) {
                    const advisorObj = teachers.find(t => 
                        (t.teacher_id || t['teacher id']) === advisorId ||
                        (t.teacher_name || t['teacher name']) === advisorId
                    );
                    const advisorName = advisorObj ? (advisorObj.teacher_name || advisorObj['teacher name']) : advisorId;
                    
                    advisorSpan = `<span class="advisor-inline"> ( <strong>ครูที่ปรึกษา:</strong> ${advisorName} )</span>`;
                }
            }
        }
        else if(viewType === 'teacher') {
            const t = teachers.find(t => (t.teacher_id || t['teacher id']) === viewId);
            title = `ตารางสอน: ${t ? (t.teacher_name || t['teacher name']) : viewId}`;
        }
        else if(viewType === 'room') {
            const r = rooms.find(r => (r.room_id || r['room id']) === viewId);
            title = `ตารางการใช้ห้อง: ${r ? (r.room_name || r['room name']) : viewId}`;
        }

        html += `<div class="header-title-box">
            <h2>${title}</h2>
            ${advisorSpan}
        </div>
        
        <table class="timetable">
            <thead>
                <tr>
                    <th style="width:80px;">วัน/เวลา</th>
                    ${periods.map(p => `
                        <th>
                            <span class="time-header">${getPeriodTime(p)}</span>
                            <span class="period-label">(คาบ ${p})</span>
                        </th>
                    `).join('')}
                </tr>
            </thead>
            <tbody>`;

        html += daysEn.map(day => {
            let cols = periods.map(p => {
                if (p === 5) return `<td class="break-col">พัก</td>`;
                const slot = timeslots.find(ts => ts.day === day && parseInt(ts.period) === p);
                const tSlotId = slot?.timeslot_id || slot?.['timeslot id'];
                
                const entry = timetable.find(t => {
                    if (t.timeslot_id !== tSlotId) return false;
                    if (viewType === 'group') return t.group_id === viewId;
                    if (viewType === 'teacher') return t.teacher_id === viewId;
                    if (viewType === 'room') return t.room_id === viewId;
                    return false;
                });

                if (entry) {
                    const t = teachers.find(tc => (tc.teacher_id || tc['teacher id']) === entry.teacher_id);
                    const r = rooms.find(rm => (rm.room_id || rm['room id']) === entry.room_id);
                    const g = groups.find(gp => (gp.group_id || gp['group id']) === entry.group_id);

                    let infoHTML = '';
                    let tName = t ? (t.teacher_name || t['teacher name']) : entry.teacher_id;
                    if(tName && tName.startsWith('ครู')) tName = tName.replace('ครู','');

                    const rName = r ? (r.room_name || r['room name']) : entry.room_id;
                    
                    if (viewType === 'group') {
                        infoHTML = `<span class="cell-info txt-teacher">${tName}</span>
                                    <span class="cell-info txt-room">ห้อง ${rName}</span>`;
                    } else if (viewType === 'teacher') {
                        infoHTML = `<span class="cell-info txt-group">กลุ่ม ${entry.group_id}</span>
                                    <span class="cell-info txt-room">ห้อง ${rName}</span>`;
                    } else if (viewType === 'room') {
                        infoHTML = `<span class="cell-info txt-teacher">${tName}</span>
                                    <span class="cell-info txt-group">กลุ่ม ${entry.group_id}</span>`;
                    }

                    return `<td style="background:#e3f2fd;">
                        <div class="cell-content">
                            <span class="cell-sub">${entry.subject_id}</span>
                            ${infoHTML}
                        </div>
                    </td>`;
                }
                return `<td></td>`;
            }).join('');
            return `<tr><td class="day-col">${dayMap[day]}</td>${cols}</tr>`;
        }).join('');
        html += `</tbody></table>`;

        const relevantEntries = timetable.filter(t => {
            if (viewType === 'group') return t.group_id === viewId;
            if (viewType === 'teacher') return t.teacher_id === viewId;
            if (viewType === 'room') return t.room_id === viewId;
            return false;
        });

        // Group by "SubjectID"
        const groupedData = {};
        relevantEntries.forEach(entry => {
            const key = entry.subject_id; 
            if(!groupedData[key]) {
                groupedData[key] = { subject_id: entry.subject_id, count: 0 };
            }
            groupedData[key].count++; 
        });

        let grandTotalTheory = 0, grandTotalPractice = 0, grandTotalCredit = 0, grandTotalPeriods = 0;
        
        let allRows = Object.values(groupedData).map((item, index) => {
            const sub = subjects.find(s => (s.subject_id || s['subject id']) === item.subject_id);
            if(sub) {
                const t = parseInt(sub.theory || 0);
                const p = parseInt(sub.practice || 0);
                const c = parseInt(sub.credit || 0);
                const actual_periods = item.count; 

                grandTotalTheory += t;
                grandTotalPractice += p;
                grandTotalCredit += c;
                grandTotalPeriods += actual_periods;

                return { 
                    index: index+1, 
                    id: item.subject_id, 
                    name: sub.subject_name,
                    t, p, c, 
                    total: actual_periods 
                };
            }
            return null;
        }).filter(x => x !== null);

        const midPoint = Math.ceil(allRows.length / 2);
        const leftRows = allRows.slice(0, midPoint);
        const rightRows = allRows.slice(midPoint);

        const calculateSum = (rows) => rows.reduce((acc, curr) => ({
            t: acc.t + curr.t, 
            p: acc.p + curr.p, 
            c: acc.c + curr.c, 
            total: acc.total + curr.total
        }), { t:0, p:0, c:0, total:0 });

        const leftSum = calculateSum(leftRows);
        const rightSum = calculateSum(rightRows);

        const tableHeader = `
            <thead>
                <tr>
                    <th style="width:30px;">ลำดับ</th>
                    <th style="width:80px;">รหัสวิชา</th>
                    <th>ชื่อวิชา</th>
                    <th style="width:30px;">ท.</th>
                    <th style="width:30px;">ป.</th>
                    <th style="width:30px;">นก.</th>
                    <th style="width:30px;">รวม</th>
                </tr>
            </thead>`;
            
        const renderTableBody = (rows) => rows.map(r => `
            <tr>
                <td>${r.index}</td>
                <td style="text-align:left;">${r.id}</td>
                <td style="text-align:left;">${r.name}</td>
                <td>${r.t}</td>
                <td>${r.p}</td>
                <td>${r.c}</td>
                <td>${r.total}</td>
            </tr>
        `).join('');

        const renderTableFooter = (sum) => `
            <tfoot>
                <tr>
                    <td colspan="3" style="text-align:right;">รวมย่อย</td>
                    <td>${sum.t}</td>
                    <td>${sum.p}</td>
                    <td>${sum.c}</td>
                    <td>${sum.total}</td>
                </tr>
            </tfoot>`;

        const uniqueGroupIds = [...new Set(relevantEntries.map(e => e.group_id))].sort();
        let groupLegendHTML = '';
        if (uniqueGroupIds.length > 0) {
            groupLegendHTML = `<div class="group-legend">
                <h4>คำอธิบายรหัสกลุ่มเรียน</h4>
                <ul>`;
            uniqueGroupIds.forEach(gId => {
                const group = groups.find(g => (g.group_id || g['group id']) === gId);
                const gName = group ? (group.group_name || group['group name']) : 'Unknown';
                groupLegendHTML += `<li><strong>${gId}</strong> : ${gName}</li>`;
            });
            groupLegendHTML += `</ul></div>`;
        }

        let errorLogHTML = '';
        if (unassignedLog.length > 0) {
            const filteredLog = unassignedLog.filter(l => {
                if (viewType === 'group') return l.group_id === viewId;
                if (viewType === 'teacher') return l.teacher_id === viewId;
                return true;
            });

            if (filteredLog.length > 0) {
                errorLogHTML = `<div class="error-log">
                    <h3>⚠️ รายวิชาที่จัดลงตารางไม่ได้ (No Constraint - System Full)</h3>
                    <ul>
                        ${filteredLog.map(l => `<li><strong>${l.subject_id}</strong> (${l.subject_name}) - ขาด ${l.missing} คาบ [${l.reason}]</li>`).join('')}
                    </ul>
                </div>`;
            }
        }

        html += `
        <div class="summary-container">
            <table class="summary" style="width:50%;">${tableHeader}<tbody>${renderTableBody(leftRows)}</tbody>${renderTableFooter(leftSum)}</table>
            <table class="summary" style="width:50%;">${tableHeader}<tbody>${renderTableBody(rightRows)}</tbody>${renderTableFooter(rightSum)}</table>
        </div>
        <div class="grand-total-box">
            <span class="total-item"><b>รวมภาพรวมทั้งหมด:</b></span>
            <span class="total-item">ทฤษฏี: <b>${grandTotalTheory}</b></span>
            <span class="total-item">ปฏิบัติ: <b>${grandTotalPractice}</b></span>
            <span class="total-item">หน่วยกิต: <b>${grandTotalCredit}</b></span>
            <span class="total-item">รวมคาบสอนจริง: <b>${grandTotalPeriods} คาบ</b></span>
        </div>
        ${groupLegendHTML}
        ${errorLogHTML}

        <div class="pdf-footer">
            <div class="signature-box">
                <div>ตรวจแล้ว</div>
                <div class="signature-line"></div>
                <div>(.......................................................)</div>
                <div class="signature-role">หัวหน้าแผนกวิชา</div>
            </div>
            <div class="signature-box">
                <div>ตรวจแล้ว</div>
                <div class="signature-line"></div>
                <div>(.......................................................)</div>
                <div class="signature-role">หัวหน้างานพัฒนาหลักสูตรฯ</div>
            </div>
            <div class="signature-box">
                <div>เห็นควรอนุญาต</div>
                <div class="signature-line"></div>
                <div>(.......................................................)</div>
                <div class="signature-role">รองผู้อำนวยการฝ่ายวิชาการ</div>
            </div>
            <div class="signature-box">
                <div>อนุญาต</div>
                <div class="signature-line"></div>
                <div>(.......................................................)</div>
                <div class="signature-role">ผู้อำนวยการวิทยาลัย...</div>
            </div>
        </div>
        `;
    }

    html += `</div></body></html>`;
    return html;
}

app.get('/', async (req, res) => {
    try {
        const [teachers, rooms, subjects, timeslots, registers, teachAssignments, groups] = await Promise.all([
            readCSV('teacher.csv'), readCSV('room.csv'), readCSV('subject.csv'),
            readCSV('timeslot.csv'), readCSV('register.csv'), readCSV('teach.csv'), readCSV('student_group.csv')
        ]);

        const timetableResult = solveTimetable(teachers, rooms, subjects, timeslots, registers, teachAssignments, groups);
        
        const header = 'group_id,timeslot_id,subject_id,teacher_id,room_id\n';
        const rows = timetableResult.timetable.map(d => `${d.group_id},${d.timeslot_id},${d.subject_id},${d.teacher_id},${d.room_id}`).join('\n');
        fs.writeFileSync('output.csv', '\ufeff' + header + rows, 'utf8');

        const { type, id } = req.query;
        const html = renderHTML(type, id, timetableResult, timeslots, subjects, teachers, rooms, groups);
        res.send(html);

    } catch (e) {
        console.error(e);
        res.status(500).send("Server Error: " + e.message);
    }
});

app.get('/download-csv', (req, res) => {
    const file = path.join(__dirname, 'output.csv');
    if(fs.existsSync(file)) res.download(file);
    else res.status(404).send("File not found");
});

app.get('/download-excel', async (req, res) => {
    try {
        const { type, id } = req.query;
        if(!type || !id) return res.send('Please select a schedule first.');

        const [teachers, rooms, subjects, timeslots, registers, teachAssignments, groups] = await Promise.all([
            readCSV('teacher.csv'), readCSV('room.csv'), readCSV('subject.csv'),
            readCSV('timeslot.csv'), readCSV('register.csv'), readCSV('teach.csv'), readCSV('student_group.csv')
        ]);
        const { timetable } = solveTimetable(teachers, rooms, subjects, timeslots, registers, teachAssignments, groups);

        const filteredTimetable = timetable.filter(t => {
            if (type === 'group') return t.group_id === id;
            if (type === 'teacher') return t.teacher_id === id;
            if (type === 'room') return t.room_id === id;
            return false;
        });

        const workbook = new ExcelJS.Workbook();
        const sheet = workbook.addWorksheet('Timetable');

        sheet.columns = [
            { header: 'วัน/คาบ', key: 'day', width: 15 },
            ...Array.from({length:12}, (_, i) => ({ header: `คาบ ${i+1}`, key: `p${i+1}`, width: 12 }))
        ];

        const daysEn = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
        const dayMap = { 'Mon': 'วันจันทร์', 'Tue': 'วันอังคาร', 'Wed': 'วันพุธ', 'Thu': 'วันพฤหัสบดี', 'Fri': 'วันศุกร์' };

        daysEn.forEach(day => {
            let rowData = { day: dayMap[day] };
            for(let p=1; p<=12; p++) {
                if(p===5) {
                    rowData[`p${p}`] = 'พัก';
                    continue;
                }
                const slot = timeslots.find(ts => ts.day === day && parseInt(ts.period) === p);
                const tSlotId = slot?.timeslot_id || slot?.['timeslot id'];
                const entry = filteredTimetable.find(t => t.timeslot_id === tSlotId);
                
                if(entry) {
                    rowData[`p${p}`] = `${entry.subject_id}\n(${entry.room_id})`;
                } else {
                    rowData[`p${p}`] = '';
                }
            }
            const row = sheet.addRow(rowData);
            row.height = 40;
            row.eachCell((cell) => {
                cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
                cell.border = { top: {style:'thin'}, left: {style:'thin'}, bottom: {style:'thin'}, right: {style:'thin'} };
            });
        });

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