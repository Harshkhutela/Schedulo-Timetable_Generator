const express = require('express');
const router = express.Router();
const Timetable = require('../models/Timetable');
const Subject = require('../models/Subject');
const Teacher = require('../models/Teacher');

router.get('/about', (req, res) => {
  res.render('about');
});

// Middleware to ensure user is authenticated
function ensureAuthenticated(req, res, next) {
  if (req.isAuthenticated()) return next();
  res.redirect('/auth/login');
}

// 🔁 Add this helper function
function getFullDetails(subjectCode, subjectMap) {
  if (!subjectCode || !subjectMap) return { fullSubject: subjectCode, fullTeacher: "" };
  const entry = subjectMap.find(item => item.code === subjectCode);
  return {
    fullSubject: entry ? entry.fullName : subjectCode,
    fullTeacher: entry ? entry.teacher : ""
  };
}

// ✅ Full Timetable View (unchanged)
router.get('/timetable', ensureAuthenticated, async (req, res) => {
  try {
    const selectedCourse = req.query.course || null;
    const latestTimetableDoc = await Timetable.findOne().sort({ createdAt: -1 });

    if (!latestTimetableDoc) {
      return res.render('user/timetable', {
        timetable: null,
        courses: [],
        selectedCourse: null
      });
    }

    const {
      timetable: fullTimetable,
      subjectTeachers,
      university,
      faculty,
      effectiveFrom: wefDate,
      days,
      slots
    } = latestTimetableDoc;

    const courses = Object.keys(fullTimetable);

    let filteredTimetable = fullTimetable;
    let filteredSubjectTeachers = subjectTeachers;

    if (selectedCourse && fullTimetable[selectedCourse]) {
      filteredTimetable = { [selectedCourse]: fullTimetable[selectedCourse] };
      filteredSubjectTeachers = { [selectedCourse]: subjectTeachers[selectedCourse] || [] };
    }

    res.render('user/timetable', {
      timetable: filteredTimetable,
      university,
      faculty,
      wefDate,
      subjectTeachers: filteredSubjectTeachers,
      slots,
      days,
      courses,
      selectedCourse,
      userEmail: req.user?.email || ''
    });
  } catch (err) {
    console.error('Error fetching timetable:', err);
    res.status(500).send('Error loading timetable');
  }
});

// ✅ Today's Timetable View (updated to show full names)
router.get('/timetable/today', ensureAuthenticated, async (req, res) => {
  try {
    const selectedCourse = req.query.course || null;
    const latestTimetableDoc = await Timetable.findOne().sort({ createdAt: -1 });

    const daysMap = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const todayDate = new Date();
    const today = daysMap[todayDate.getDay()];
    const dayName = today;

    if (!latestTimetableDoc) {
      return res.render('user/today', {
        today,
        todaySchedule: [],
        slots: [],
        courses: [],
        selectedCourse: null,
        dayName,
        subjectTeachers: {}
      });
    }

    const { timetable: fullTimetable, slots, days } = latestTimetableDoc;
    const courses = Object.keys(fullTimetable);
    let todaySchedule = [];

    // 👇 Function to clean shortName
    const extractShortName = (raw) => {
      return raw?.split('(')[0]?.trim() || '';
    };

    if (selectedCourse && fullTimetable[selectedCourse]) {
      const courseSchedule = fullTimetable[selectedCourse];
      const todayIndex = days.indexOf(today);

      if (todayIndex !== -1) {
        const rawSchedule = courseSchedule[today] || [];

        let mergedSchedule = [];
        let i = 0;

        while (i < rawSchedule.length) {
          let current = rawSchedule[i];
          let startIndex = i;
          let endIndex = i;

          // Merge consecutive same entries
          while (
            endIndex + 1 < rawSchedule.length &&
            rawSchedule[endIndex + 1]?.subject === current.subject &&
            rawSchedule[endIndex + 1]?.teacher === current.teacher &&
            rawSchedule[endIndex + 1]?.room === current.room
          ) {
            endIndex++;
          }

          // 🔍 Lookup using cleaned subject shortName
          const cleanedShortName = extractShortName(current.subject);
          const subjectDoc = await Subject.findOne({ shortName: new RegExp(`^${cleanedShortName}$`, 'i') }).populate('assignedTeacher');

          const subjectFullName = subjectDoc?.fullName || current.subject || "Free";
          const teacherFullName = subjectDoc?.assignedTeacher?.name || current.teacher || "N/A";
          const isFree = subjectFullName === "Free";
          const isLab = subjectFullName.toLowerCase().includes('lab');

          const timeRange = `${slots[startIndex].split("-")[0]} - ${slots[endIndex].split("-")[1]}`;

          mergedSchedule.push({
            subject: subjectFullName,
            teacher: isFree ? "" : teacherFullName,
            room: current.room || "N/A",
            time: timeRange,
            type: isFree ? "Free" : (isLab ? "Lab" : "Lecture")
          });

          i = endIndex + 1;
        }

        todaySchedule = mergedSchedule;
      }
    }

    res.render('user/today', {
      today,
      todaySchedule,
      slots,
      courses,
      selectedCourse,
      dayName,
      subjectTeachers: {}
    });

  } catch (err) {
    console.error('Error loading today view:', err);
    res.status(500).send('Error loading today timetable');
  }
});


// ✅ Excel Export Route
const generateStyledTimetableExcel = require('../utils/excelGenerator');

router.get('/timetable/export/excel', async (req, res) => {
  try {
    const latestTimetableDoc = await Timetable.findOne().sort({ createdAt: -1 });
    if (!latestTimetableDoc) return res.status(404).send('No timetable found.');

    const workbook = generateStyledTimetableExcel(latestTimetableDoc);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=timetable.xlsx');

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Excel export error:', err);
    res.status(500).send('Error generating Excel.');
  }
});

module.exports = router;
