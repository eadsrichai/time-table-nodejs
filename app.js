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
                Object.keys(data).forEach(key => { cleanData[key.trim()] = data[key] ? data[key].trim() : ''; });
                results.push(cleanData);
            })
            .on('end', () => resolve(results))
            .on('error', (err) => resolve([]));
    });
}

// 2. Logic จัดตาราง
function solveTimetable(teachers, rooms, subjects, timeslots, registers, teachAssignments, groups) {
    let timetable = [];
    let groupDailyLoad = {};

    const getGroupLoad = (groupId, day) => {
        if (!groupDailyLoad[groupId]) groupDailyLoad[groupId] = { Mon:0, Tue:0, Wed:0, Thu:0, Fri:0 };
        return groupDailyLoad[groupId][day] || 0;
    };

    const incrementGroupLoad = (groupId, day, amount = 1) => {
        if (!groupDailyLoad[groupId]) groupDailyLoad[groupId] = { Mon:0, Tue:0, Wed:0, Thu:0, Fri:0 };
        groupDailyLoad[groupId][day] += amount;
    };

    const isSlotFree = (tSlotId, teacherId, groupId, roomId) => {
        const teacherBusy = timetable.find(t => t.timeslot_id === tSlotId && t.teacher_id === teacherId);
        if (teacherBusy) return false;
        const groupBusy = timetable.find(t => t.timeslot_id === tSlotId && t.group_id === groupId);
        if (groupBusy) return false;
        if (roomId) {
            const roomBusy = timetable.find(t => t.timeslot_id === tSlotId && t.room_id === roomId);
            if (roomBusy) return false;
        }
        return true;
    };

    const findTargetRoomType = (subjectName, groupName, isPractice) => {
        const sName = subjectName ? subjectName.toLowerCase() : '';
        const gName = groupName ? groupName.toLowerCase() : '';

        if (gName.includes('ทวิ') || gName.includes('dual') || gName.includes('dve')) {
            return 'Factory';
        }
        
        if (sName.includes('iot') || sName.includes('internet of things')) return 'IOT LAB';
        if (sName.includes('network') || sName.includes('เครือข่าย')) return 'Network Lab';
        if (sName.includes('ai ') || sName.includes('intelligence') || sName.includes('ปัญญาประดิษฐ์') || sName.includes('วิทยาการข้อมูล')) return 'AI Lab';
        if (sName.includes('graphic') || sName.includes('กราฟิก') || sName.includes('3d') || sName.includes('multimedia') || sName.includes('game')) return 'Computer Graphic Lab';
        if (sName.includes('computer') || sName.includes('คอมพิวเตอร์') || sName.includes('programming') || sName.includes('database')) return 'Computer Lab';

        if (isPractice) return 'Practice_General'; 
        return 'Theory'; 
    };

    const validRegisters = registers.filter(r => {
        const sId = r.subject_id || r['subject id'];
        return subjects.some(s => (s.subject_id || s['subject id']) === sId);
    });

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
        const subject = subjects.find(s => (s.subject_id || s['subject id']) === sId);
        const groupObj = groups.find(g => (g.group_id || g['group id']) === gId);
        
        if (!subject) continue;

        const theory = parseInt(subject.theory || 0);
        const practice = parseInt(subject.practice || 0);
        const totalPeriodsNeeded = theory + practice;
        const isPractice = practice > 0;
        
        const assignment = teachAssignments.find(a => (a.subject_id || a['subject id']) === sId);
        const teacherId = assignment ? (assignment.teacher_id || assignment['teacher id']) : null;
        const teacherInfo = teachers.find(t => (t.teacher_id || t['teacher id']) === teacherId);

        const targetRoomType = findTargetRoomType(
            subject.subject_name, 
            groupObj ? groupObj.group_name : '', 
            isPractice
        );

        let isBlockBooked = false;

        if (totalPeriodsNeeded > 1) {
            let blockCandidates = [];
            const days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
            
            for (const day of days) {
                for (let startP = 1; startP <= 12 - totalPeriodsNeeded + 1; startP++) {
                    let isValidBlock = true;
                    let blockSlots = [];

                    for (let k = 0; k < totalPeriodsNeeded; k++) {
                        const currentP = startP + k;
                        if (currentP === 5) { isValidBlock = false; break; }
                        if (teacherInfo?.role === 'Leader' && day === 'Tue' && currentP === 8) { isValidBlock = false; break; }

                        const slotObj = timeslots.find(ts => ts.day === day && parseInt(ts.period) === currentP);
                        if (!slotObj) { isValidBlock = false; break; }

                        if (!isSlotFree(slotObj.timeslot_id || slotObj['timeslot id'], teacherId, gId, null)) {
                            isValidBlock = false; break;
                        }
                        blockSlots.push(slotObj);
                    }

                    if (isValidBlock) {
                        const validRoom = rooms.find(r => {
                            const rType = r.room_type ? r.room_type.trim() : '';
                            const rName = r.room_name ? r.room_name.toLowerCase() : '';
                            let isTypeMatch = false;

                            if (targetRoomType === 'Factory') isTypeMatch = (rType === 'Factory');
                            else if (targetRoomType === 'Theory') isTypeMatch = (rType === 'Theory');
                            else if (targetRoomType === 'Practice_General') isTypeMatch = (rType !== 'Theory') && (rType !== 'Factory');
                            else {
                                const targetKeyword = targetRoomType.toLowerCase();
                                isTypeMatch = (rType.toLowerCase() === targetKeyword) || rName.includes(targetKeyword.replace(' lab','')); 
                            }

                            if (!isTypeMatch) return false;
                            return blockSlots.every(slot => {
                                return isSlotFree(slot.timeslot_id || slot['timeslot id'], null, null, r.room_id || r['room id']);
                            });
                        });

                        if (validRoom) {
                            let timePenalty = (startP + totalPeriodsNeeded > 10) ? 500 : 0;
                            let loadPenalty = getGroupLoad(gId, day) * 10; 
                            let priorityScore = timePenalty + loadPenalty + startP;
                            blockCandidates.push({ score: priorityScore, slots: blockSlots, room: validRoom, day: day });
                        }
                    }
                }
            }

            if (blockCandidates.length > 0) {
                blockCandidates.sort((a, b) => a.score - b.score);
                const bestBlock = blockCandidates[0];
                bestBlock.slots.forEach(slot => {
                    timetable.push({
                        group_id: gId, timeslot_id: slot.timeslot_id || slot['timeslot id'],
                        subject_id: sId, teacher_id: teacherId, room_id: bestBlock.room.room_id || bestBlock.room['room id']
                    });
                });
                incrementGroupLoad(gId, bestBlock.day, totalPeriodsNeeded);
                isBlockBooked = true;
            }
        }

        if (!isBlockBooked) {
            for (let i = 0; i < totalPeriodsNeeded; i++) {
                let candidates = [];
                for (const slot of timeslots) {
                    const tSlotId = slot.timeslot_id || slot['timeslot id'];
                    const period = parseInt(slot.period);
                    const day = slot.day;

                    if (period === 5) continue;
                    if (teacherInfo?.role === 'Leader' && day === 'Tue' && period === 8) continue;
                    if (!isSlotFree(tSlotId, teacherId, gId, null)) continue;

                    const room = rooms.find(r => {
                        const rId = r.room_id || r['room id'];
                        const rType = r.room_type ? r.room_type.trim() : '';
                        const rName = r.room_name ? r.room_name.toLowerCase() : '';
                        let isTypeMatch = false;
                        if (targetRoomType === 'Factory') isTypeMatch = (rType === 'Factory');
                        else if (targetRoomType === 'Theory') isTypeMatch = (rType === 'Theory');
                        else if (targetRoomType === 'Practice_General') isTypeMatch = (rType !== 'Theory') && (rType !== 'Factory');
                        else {
                            const targetKeyword = targetRoomType.toLowerCase();
                            isTypeMatch = (rType.toLowerCase() === targetKeyword) || rName.includes(targetKeyword.replace(' lab',''));
                        }
                        return isSlotFree(tSlotId, null, null, rId) && isTypeMatch;
                    });

                    if (room) {
                        let timePenalty = (period > 10) ? 1000 : 0;
                        let loadPenalty = getGroupLoad(gId, day) * 10;
                        let periodPenalty = period;
                        candidates.push({ score: timePenalty + loadPenalty + periodPenalty, slot: slot, room: room });
                    }
                }

                if (candidates.length > 0) {
                    candidates.sort((a, b) => a.score - b.score);
                    const best = candidates[0];
                    timetable.push({
                        group_id: gId, timeslot_id: best.slot.timeslot_id || best.slot['timeslot id'],
                        subject_id: sId, teacher_id: teacherId, room_id: best.room.room_id || best.room['room id']
                    });
                    incrementGroupLoad(gId, best.slot.day);
                }
            }
        }
    }
    return timetable;
}

// 3. ฟังก์ชันสร้าง HTML
function renderHTML(viewType, viewId, timetable, timeslots, subjects, teachers, rooms, groups) {
    const daysEn = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'];
    const dayMap = { 'Mon': 'วันจันทร์', 'Tue': 'วันอังคาร', 'Wed': 'วันพุธ', 'Thu': 'วันพฤหัสบดี', 'Fri': 'วันศุกร์' };
    const periods = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];
    
    const getPeriodTime = (p) => {
        const startHour = 8 + (p - 1);
        const endHour = 8 + p;
        const formatTime = (h) => (h < 10 ? '0' + h : h) + '.00';
        return `${formatTime(startHour)}-${formatTime(endHour)}`;
    };

    const groupOptions = groups.map(g => `<option value="${g.group_id}" ${viewType === 'group' && viewId === g.group_id ? 'selected' : ''}>${g.group_id} - ${g.group_name}</option>`).join('');
    const teacherOptions = teachers.map(t => `<option value="${t.teacher_id}" ${viewType === 'teacher' && viewId === t.teacher_id ? 'selected' : ''}>${t.teacher_name}</option>`).join('');
    const roomOptions = rooms.map(r => `<option value="${r.room_id}" ${viewType === 'room' && viewId === r.room_id ? 'selected' : ''}>${r.room_id} - ${r.room_name}</option>`).join('');

    const excelLink = `'/download-excel?type=${viewType || ''}&id=${viewId || ''}'`;

    let html = `
    <html>
    <head>
        <meta charset="UTF-8">
        <title>ระบบจัดตารางสอนอัตโนมัติ</title>
        <link href="https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;600&display=swap" rel="stylesheet">
        <script src="https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js"></script>
        
        <style>
            body { 
                font-family: 'Prompt', sans-serif; 
                font-weight: 300; 
                padding: 0; 
                margin: 0; 
                background-color: #f4f6f9; 
            }
            .navbar { background-color: #004a99; padding: 15px 20px; color: white; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; box-shadow: 0 2px 5px rgba(0,0,0,0.2); }
            .navbar h1 { margin: 0; font-size: 20px; font-weight: 600; display: flex; align-items: center; }
            .navbar h1 span { font-size: 12px; background: #2196F3; padding: 2px 8px; border-radius: 10px; margin-left: 10px; font-weight: 400;}
            
            .menu-container { display: flex; gap: 10px; align-items: center; }
            .menu-item { display: flex; flex-direction: column; }
            .menu-item label { font-size: 10px; margin-bottom: 2px; color: #bbdefb; }
            
            select { 
                padding: 8px; 
                border-radius: 4px; 
                border: none; 
                font-family: 'Prompt', sans-serif; 
                font-weight: 300;
                font-size: 14px; 
                min-width: 180px; 
                cursor: pointer; 
            }
            select:focus { outline: 2px solid #82b1ff; }
            
            .btn-group { display: flex; gap: 5px; margin-left: 10px; }
            .btn { 
                color: white; 
                text-decoration: none; 
                padding: 8px 12px; 
                border-radius: 4px; 
                font-size: 14px; 
                transition: 0.3s; 
                cursor: pointer; 
                border: none; 
                display: flex; 
                align-items: center; 
                gap: 5px; 
                font-family: 'Prompt', sans-serif; 
                font-weight: 300;
            }
            .btn-excel { background: #2e7d32; }
            .btn-excel:hover { background: #1b5e20; }
            .btn-pdf { background: #d32f2f; }
            .btn-pdf:hover { background: #b71c1c; }
            .btn-csv { background: #f57f17; }
            
            .content { padding: 20px; max-width: 1280px; margin: 0 auto; padding-bottom: 50px; background: white; }
            
            table.timetable { width: 100%; border-collapse: collapse; margin-top: 20px; table-layout: fixed; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
            table.timetable th, table.timetable td { border: 1px solid #000; padding: 4px; text-align: center; font-size: 11px; height: 55px; vertical-align: middle; overflow: hidden; }
            table.timetable th { background-color: #e3f2fd; color: #333; font-weight: 600; border-bottom: 2px solid #2196F3; }
            
            .summary-container { display: flex; gap: 10px; margin-top: 15px; align-items: flex-start; }
            table.summary { width: 100%; border-collapse: collapse; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
            table.summary th, table.summary td { border: 1px solid #000; padding: 6px; text-align: center; font-size: 12px; }
            table.summary th { background-color: #004a99; color: white; height: 35px; font-weight: 400; }
            table.summary tfoot td { background-color: #f1f1f1; font-weight: 600; color: #333; }
            
            .grand-total-box { background-color: #e8f5e9; border: 2px solid #4caf50; border-radius: 5px; padding: 10px; text-align: center; margin-top: 10px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); display: flex; justify-content: space-around; page-break-inside: avoid; }
            
            .group-legend { margin-top: 20px; border-top: 1px dashed #ccc; padding-top: 10px; font-size: 12px; color: #555; }
            .group-legend h4 { margin: 0 0 5px 0; color: #333; font-weight: 600; }
            .group-legend ul { list-style: none; padding: 0; margin: 0; display: flex; flex-wrap: wrap; gap: 15px; }
            .group-legend li { background: #f5f5f5; padding: 2px 8px; border-radius: 4px; border: 1px solid #eee; }
            .group-legend strong { color: #388e3c; font-weight: 600; }

            .time-header { font-size: 12px; display: block; margin-bottom: 2px; font-weight: 600; }
            .period-label { font-size: 9px; color: #666; font-weight: 300; }
            .day-col { background-color: #fafafa; font-weight: 600; width: 80px; color: #004a99; border-right: 2px solid #ddd; }
            
            .break-col { 
                background-color: #eee; 
                font-weight: 600; 
                color: #777; 
                vertical-align: middle; 
                text-align: center;
                white-space: nowrap;
            }
            
            .cell-content { display: flex; flex-direction: column; justify-content: center; height: 100%; }
            .cell-sub { font-weight: 600; font-size: 13px; color: #000; margin-bottom: 3px; }
            .cell-info { font-size: 11px; margin-top: 2px; }
            .txt-teacher { color: #d32f2f; }
            .txt-room { color: #1976d2; }
            .txt-group { color: #388e3c; }
            .welcome-box { text-align: center; margin-top: 50px; color: #666; }

            /* --- CSS สำหรับโหมด Fit to Page A4 Landscape --- */
            .fit-to-page {
                width: 100% !important;
                padding: 0 !important;
            }
            .fit-to-page h2 {
                font-size: 14px !important; /* ย่อหัวข้อ */
                margin-top: 5px !important;
                margin-bottom: 5px !important;
            }
            .fit-to-page table.timetable th, 
            .fit-to-page table.timetable td {
                font-size: 9px !important; /* ย่อตัวหนังสือในตารางเรียน */
                padding: 2px !important;
                height: 35px !important; /* ลดความสูงแถว */
            }
            .fit-to-page .cell-sub { font-size: 10px !important; margin-bottom: 1px !important; }
            .fit-to-page .cell-info { font-size: 8px !important; }
            .fit-to-page .time-header { font-size: 9px !important; }
            .fit-to-page .period-label { display: none !important; } /* ซ่อนคำว่า (คาบที่..) เพื่อประหยัดที่ */
            
            /* ย่อตารางสรุป */
            .fit-to-page .summary-container { margin-top: 5px !important; gap: 5px !important; }
            .fit-to-page table.summary th, 
            .fit-to-page table.summary td {
                font-size: 9px !important;
                padding: 2px !important;
                height: 20px !important;
            }
            
            /* ย่อกล่องยอดรวม */
            .fit-to-page .grand-total-box {
                margin-top: 5px !important;
                padding: 5px !important;
            }
            .fit-to-page .grand-total-box .total-item {
                font-size: 10px !important;
            }
            
            /* ย่อ Legend */
            .fit-to-page .group-legend {
                margin-top: 5px !important;
                padding-top: 5px !important;
                font-size: 9px !important;
            }
            .fit-to-page .group-legend h4 { margin-bottom: 2px !important; font-size: 10px !important;}
            .fit-to-page .group-legend li { padding: 1px 4px !important; }

            @media print {
                .navbar { display: none; }
                .content { padding: 0; max-width: 100%; }
                table { page-break-inside: auto; }
            }
        </style>
        <script>
            function navigate(type, id) {
                if(id) window.location.href = '/?type=' + type + '&id=' + id;
            }
            
            function exportPDF() {
                const btn = document.querySelector('.btn-pdf');
                const element = document.getElementById('report-content');
                const originalText = btn.innerText;
                
                // 1. ใส่ Class ย่อขนาด (Fit to Page)
                element.classList.add('fit-to-page');
                btn.innerText = 'กำลังสร้าง...';
                
                const opt = {
                    margin: 5, // หน่วย mm (ขอบบางๆ)
                    filename: 'timetable-report.pdf',
                    image: { type: 'jpeg', quality: 0.98 },
                    html2canvas: { scale: 2, useCORS: true }, // scale 2 เพื่อความคมชัด
                    jsPDF: { unit: 'mm', format: 'a4', orientation: 'landscape' }
                };
                
                html2pdf().set(opt).from(element).output('bloburl').then(function(pdfUrl) {
                    window.open(pdfUrl, '_blank');
                    
                    // 2. คืนค่าเดิมหลังจากสร้างเสร็จ
                    element.classList.remove('fit-to-page');
                    btn.innerText = originalText;
                }).catch(err => {
                    console.error(err);
                    alert('Error generating PDF');
                    element.classList.remove('fit-to-page');
                    btn.innerText = originalText;
                });
            }

            function downloadExcel() {
                window.location.href = ${excelLink};
            }
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
                    <button onclick="exportPDF()" class="btn btn-pdf">📕 PDF (One Page)</button>
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
        if(viewType === 'group') title = `ตารางเรียนกลุ่ม: ${groups.find(g => g.group_id === viewId)?.group_name || viewId}`;
        if(viewType === 'teacher') title = `ตารางสอน: ${teachers.find(t => t.teacher_id === viewId)?.teacher_name || viewId}`;
        if(viewType === 'room') title = `ตารางการใช้ห้อง: ${rooms.find(r => r.room_id === viewId)?.room_name || viewId}`;

        // --- Timetable Section ---
        html += `<h2 style="color:#333; border-left:5px solid #004a99; padding-left:10px; margin-top:20px;">${title}</h2>
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
                    const teacherName = t?.teacher_name.split(' ')[0] || entry.teacher_id;

                    if (viewType === 'group') {
                        infoHTML = `<span class="cell-info txt-teacher">${teacherName}</span>
                                    <span class="cell-info txt-room">ห้อง ${r?.room_name || entry.room_id}</span>`;
                    } else if (viewType === 'teacher') {
                        infoHTML = `<span class="cell-info txt-group">กลุ่ม ${g?.group_id || entry.group_id}</span>
                                    <span class="cell-info txt-room">ห้อง ${r?.room_name || entry.room_id}</span>`;
                    } else if (viewType === 'room') {
                        infoHTML = `<span class="cell-info txt-teacher">${teacherName}</span>
                                    <span class="cell-info txt-group">กลุ่ม ${g?.group_id || entry.group_id}</span>`;
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

        // --- Summary Logic (Unique Subjects Only) ---
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

        let grandTotalTheory = 0, grandTotalPractice = 0, grandTotalCredit = 0, grandTotalPeriods = 0, grandTotalStdPeriods = 0;
        
        let allRows = Object.values(groupedData).map((item, index) => {
            const sub = subjects.find(s => (s.subject_id || s['subject id']) === item.subject_id);
            if(sub) {
                const t = parseInt(sub.theory || 0);
                const p = parseInt(sub.practice || 0);
                const c = parseInt(sub.credit || 0);
                const std_periods = t + p; 
                const actual_periods = item.count; 

                grandTotalTheory += t;
                grandTotalPractice += p;
                grandTotalCredit += c;
                grandTotalStdPeriods += std_periods;
                grandTotalPeriods += actual_periods;

                return { 
                    index: index+1, 
                    id: item.subject_id, 
                    name: sub.subject_name,
                    t, p, c, std_periods, 
                    total_periods: actual_periods 
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
            std_periods: acc.std_periods + curr.std_periods,
            total_periods: acc.total_periods + curr.total_periods
        }), { t:0, p:0, c:0, std_periods:0, total_periods:0 });

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
                    <th style="width:30px;">คาบ</th>
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
                <td>${r.std_periods}</td>
                <td>${r.total_periods}</td>
            </tr>
        `).join('');

        const renderTableFooter = (sum) => `
            <tfoot>
                <tr>
                    <td colspan="3" style="text-align:right;">รวมย่อย</td>
                    <td>${sum.t}</td>
                    <td>${sum.p}</td>
                    <td>${sum.c}</td>
                    <td>${sum.std_periods}</td>
                    <td>${sum.total_periods}</td>
                </tr>
            </tfoot>`;

        const uniqueGroupIds = [...new Set(relevantEntries.map(e => e.group_id))].sort();
        let groupLegendHTML = '';
        if (uniqueGroupIds.length > 0) {
            groupLegendHTML = `<div class="group-legend">
                <h4>คำอธิบายรหัสกลุ่มเรียน</h4>
                <ul>`;
            uniqueGroupIds.forEach(gId => {
                const group = groups.find(g => g.group_id === gId);
                const gName = group ? group.group_name : 'Unknown';
                groupLegendHTML += `<li><strong>${gId}</strong> : ${gName}</li>`;
            });
            groupLegendHTML += `</ul></div>`;
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
            <span class="total-item">คาบตามหลักสูตร: <b>${grandTotalStdPeriods}</b></span>
            <span class="total-item">ชั่วโมงสอนจริง: <b>${grandTotalPeriods}</b></span>
        </div>
        ${groupLegendHTML}
        `;
    }

    html += `</div></body></html>`;
    return html;
}

// 4. Main Route
app.get('/', async (req, res) => {
    try {
        const [teachers, rooms, subjects, timeslots, registers, teachAssignments, groups] = await Promise.all([
            readCSV('teacher.csv'), readCSV('room.csv'), readCSV('subject.csv'),
            readCSV('timeslot.csv'), readCSV('register.csv'), readCSV('teach.csv'), readCSV('student_group.csv')
        ]);

        const timetable = solveTimetable(teachers, rooms, subjects, timeslots, registers, teachAssignments, groups);
        
        const header = 'group_id,timeslot_id,subject_id,teacher_id,room_id\n';
        const rows = timetable.map(d => `${d.group_id},${d.timeslot_id},${d.subject_id},${d.teacher_id},${d.room_id}`).join('\n');
        fs.writeFileSync('output.csv', '\ufeff' + header + rows, 'utf8');

        const { type, id } = req.query;
        const html = renderHTML(type, id, timetable, timeslots, subjects, teachers, rooms, groups);
        res.send(html);

    } catch (e) {
        console.error(e);
        res.status(500).send("Server Error: " + e.message);
    }
});

// 5. Route Download CSV
app.get('/download-csv', (req, res) => {
    const file = path.join(__dirname, 'output.csv');
    if(fs.existsSync(file)) res.download(file);
    else res.status(404).send("File not found");
});

// 6. Route Download Excel
app.get('/download-excel', async (req, res) => {
    try {
        const { type, id } = req.query;
        if(!type || !id) return res.send('Please select a schedule first.');

        const [teachers, rooms, subjects, timeslots, registers, teachAssignments, groups] = await Promise.all([
            readCSV('teacher.csv'), readCSV('room.csv'), readCSV('subject.csv'),
            readCSV('timeslot.csv'), readCSV('register.csv'), readCSV('teach.csv'), readCSV('student_group.csv')
        ]);
        const timetable = solveTimetable(teachers, rooms, subjects, timeslots, registers, teachAssignments, groups);

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
    console.log(`\n[SUCCESS] Server Running at http://203.150.107.92:${PORT}`);
});
