function timetableGenerator(courses, subjects, teachers, classrooms) {
  const timetable = {};    
  const subjectTeachers = {};
  const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday'];
  const slots = [
    '9:00-10:00', '10:00-11:00', '11:00-12:00', '12:00-1:00',
    '1:00-1:30',
    '1:30-2:30', '2:30-3:30', '3:30-4:30', '4:30-5:30'
  ];

  const totalSlots = slots.length;
  const totalDays = days.length;

  const roomSchedule = {};
  const teacherSchedule = {};

  days.forEach(day => {
    roomSchedule[day] = Array(totalSlots).fill().map(() => []);
    teacherSchedule[day] = Array(totalSlots).fill().map(() => []);
  });

  const firstHalfSlots = [0, 1, 2, 3];
  const secondHalfSlots = [5, 6, 7, 8]; // skip index 4 (lunch)

  for (const course of courses) {
    const courseName = course.courseShortName;
    timetable[courseName] = {};
    subjectTeachers[courseName] = [];

    days.forEach(day => {
      timetable[courseName][day] = Array(totalSlots).fill(null);
      timetable[courseName][day][4] = {
        room: '',
        subject: 'Break',
        teacher: ''
      }; // fixed lunch break
    });

    const courseSubjects = course.subjects;
    const shuffledSubjects = [...courseSubjects].sort(() => Math.random() - 0.5);

    let assignedRoomForCourse = null;

    for (const subject of shuffledSubjects) {
      const teacher = teachers.find(t =>
        t.assignments.some(a =>
          String(a.course?._id) === String(course._id) &&
          String(a.subject?._id) === String(subject._id)
        )
      ) || subject.assignedTeacher || null;

      subjectTeachers[courseName].push({
        subjectShort: subject.shortName,
        subjectLong: subject.fullName,
        teacherShort: teacher?.shortName || 'TBA',
        teacherLong: teacher?.name || 'To Be Assigned'
      });

      const isLab = subject.isLab;
      const requiredSessions = subject.count;
      let scheduled = 0;
      const perDayPlacement = {};

      while (scheduled < requiredSessions) {
        const availableDays = days.filter(day => (perDayPlacement[day] || 0) < 2);
        if (availableDays.length === 0) break;

        const shuffledDays = [...availableDays].sort(() => Math.random() - 0.5);

        let placed = false;

        for (const day of shuffledDays) {
          if (isLab) {
            const labSlotPairs = [
              [0, 1],
              [2, 3],
              [5, 6],
              [7, 8]
            ];

            let labPlaced = false;

            for (const [slot1, slot2] of labSlotPairs) {
              if (
                timetable[courseName][day][slot1] === null &&
                timetable[courseName][day][slot2] === null
              ) {
                let room = null;

                if (assignedRoomForCourse) {
                  const sameRoom = classrooms.find(r =>
                    r.classOrLab === 'lab' &&
                    r.roomNumber === assignedRoomForCourse &&
                    (!subject.labType || r.labType === subject.labType)
                  );
                  if (
                    sameRoom &&
                    !roomSchedule[day][slot1].includes(sameRoom.roomNumber) &&
                    !roomSchedule[day][slot2].includes(sameRoom.roomNumber)
                  ) {
                    room = sameRoom;
                  }
                }

                if (!room) {
                  room = classrooms.find(r =>
                    r.classOrLab === 'lab' &&
                    (!subject.labType || r.labType === subject.labType) &&
                    !roomSchedule[day][slot1].includes(r.roomNumber) &&
                    !roomSchedule[day][slot2].includes(r.roomNumber)
                  );
                }

                const teacherFree = !teacher || (
                  !teacherSchedule[day][slot1].includes(teacher.shortName) &&
                  !teacherSchedule[day][slot2].includes(teacher.shortName)
                );

                if (room && teacherFree) {
                  if (!assignedRoomForCourse) assignedRoomForCourse = room.roomNumber;

                  const labGroup = scheduled < 2 ? 'Lab-1 (G1)' : 'Lab-2 (G2)';
                  const label = `${subject.shortName} (${labGroup})`;

                  const cell = {
                    room: `${room.roomNumber}, ${room.building}`,
                    subject: label,
                    teacher: teacher?.shortName || 'TBA'
                  };

                  timetable[courseName][day][slot1] = cell;
                  timetable[courseName][day][slot2] = cell;

                  roomSchedule[day][slot1].push(room.roomNumber);
                  roomSchedule[day][slot2].push(room.roomNumber);

                  if (teacher) {
                    teacherSchedule[day][slot1].push(teacher.shortName);
                    teacherSchedule[day][slot2].push(teacher.shortName);
                  }

                  perDayPlacement[day] = (perDayPlacement[day] || 0) + 2;
                  scheduled += 2;
                  labPlaced = true;
                  break;
                }
              }
            }

            if (labPlaced) {
              placed = true;
              break;
            }
          } else {
            const trySlots = [...firstHalfSlots, ...secondHalfSlots];
            for (let i = 0; i < trySlots.length; i++) {
              const slot = trySlots[i];

              if (timetable[courseName][day][slot] === null) {
                let room = null;

                if (assignedRoomForCourse) {
                  const sameRoom = classrooms.find(r =>
                    r.classOrLab === 'class' &&
                    r.roomNumber === assignedRoomForCourse
                  );
                  if (sameRoom && !roomSchedule[day][slot].includes(sameRoom.roomNumber)) {
                    room = sameRoom;
                  }
                }

                if (!room) {
                  room = classrooms.find(r =>
                    r.classOrLab === 'class' &&
                    !roomSchedule[day][slot].includes(r.roomNumber)
                  );
                }

                const teacherFree = !teacher || !teacherSchedule[day][slot].includes(teacher.shortName);

                if (room && teacherFree) {
                  if (!assignedRoomForCourse) assignedRoomForCourse = room.roomNumber;

                  const cell = {
                    room: `${room.roomNumber}, ${room.building}`,
                    subject: subject.shortName,
                    teacher: teacher?.shortName || 'TBA'
                  };

                  timetable[courseName][day][slot] = cell;
                  roomSchedule[day][slot].push(room.roomNumber);

                  if (teacher) teacherSchedule[day][slot].push(teacher.shortName);

                  perDayPlacement[day] = (perDayPlacement[day] || 0) + 1;
                  scheduled++;
                  placed = true;
                  break;
                }
              }
            }

            if (placed) break;
          }
        }

        if (!placed) break;
      }
    }

    days.forEach((day, dIndex) => {
      timetable[courseName][day] = timetable[courseName][day].map((cell, sIndex) => {
        if (sIndex === 4) return cell; // skip lunch slot
        if (!cell && (sIndex + dIndex) % 2 === 0) {
          return { room: 'Library', subject: 'Library', teacher: '' };
        }
        return cell || { room: 'Library', subject: 'Library', teacher: '' };
      });
    });
  }

  return { timetable, subjectTeachers, days, slots };
}

module.exports = timetableGenerator;
