const express = require('express');
const router = express.Router();
const Course = require('../models/Course');
const Subject = require('../models/Subject');
const Classroom = require('../models/classroom');

// Step 3: Get Course Details
router.get('/', async (req, res) => {
  try {
    const courses = await Course.find().populate('subjects');
    const subjects = await Subject.find();
    const classrooms = await Classroom.find();
    res.render('steps/step3', { courses, subjects, classrooms });
  } catch (error) {
    console.error(error);
    res.status(500).send('Server error');
  }
});

// Step 3: Add Course
router.post('/add', async (req, res) => {
    let { courseName, courseShortName, subjects, studentStrength } = req.body;
    studentStrength = parseInt(studentStrength) || 0;
    
    // Ensure subjects is always an array
    if (!Array.isArray(subjects)) {
      if (subjects) subjects = [subjects];
      else subjects = []; // no subject selected
    }

    try {
      const newCourse = new Course({
        courseName,
        courseShortName,
        subjects,
        studentStrength
      });

      await newCourse.save();
      res.redirect('/step3');
    } catch (error) {
      console.error(error);
      res.status(500).send('Error adding course');
    }   
  });
  

// Step 3: Edit Course
router.post('/edit/:id', async (req, res) => {
  const { courseName, courseShortName, studentStrength, subjects } = req.body;

  try {
    const course = await Course.findById(req.params.id); // This line was probably missing
    if (!course) return res.status(404).send("Course not found");

    course.courseName = courseName;
    course.courseShortName = courseShortName;
    course.studentStrength = studentStrength;
    course.subjects = Array.isArray(subjects) ? subjects : [subjects];

    await course.save();
    res.redirect('/step3');
  } catch (error) {
    console.error(error);
    res.status(500).send('Error editing course');
  }
});



// Step 3: Delete Course
router.post('/delete/:id', async (req, res) => {
  try {
    await Course.findByIdAndDelete(req.params.id);
    res.redirect('/step3');
  } catch (error) {
    console.error(error);
    res.status(500).send('Error deleting course');
  }
});

module.exports = router;
