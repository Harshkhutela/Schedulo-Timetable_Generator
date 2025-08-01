const express = require('express');
const router = express.Router();
const Classroom = require('../models/classroom');

// View Step 4
router.get('/', async (req, res) => {
  try {
    const classrooms = await Classroom.find();
    res.render('steps/step4', { classrooms });
  } catch (error) {
    console.error(error);
    res.status(500).send('Server Error');
  }
});

// Add Classroom or Lab
router.post('/add-classroom', async (req, res) => {
  const { classOrLab, roomNumber, building, capacityRange, labType } = req.body;

  try {
    const newRoom = new Classroom({
      classOrLab,
      roomNumber,
      building,
      capacityRange: classOrLab === 'class' ? capacityRange : undefined,
      labType: classOrLab === 'lab' ? labType : ''
    });

    await newRoom.save();
    res.redirect('/step4');
  } catch (error) {
    console.error(error);
    res.status(500).send('Error adding room/lab');
  }
});

// Edit Room/Lab
router.post('/edit/:id', async (req, res) => {
  const { classOrLab, roomNumber, building, capacityRange, labType } = req.body;

  try {
    const room = await Classroom.findById(req.params.id);
    if (!room) return res.status(404).send('Room not found');

    room.classOrLab = classOrLab;
    room.roomNumber = roomNumber;
    room.building = building;

    if (classOrLab === 'class') {
      room.capacityRange = capacityRange;
      room.labType = '';
    } else {
      room.capacityRange = undefined;
      room.labType = labType;
    }

    await room.save();
    res.redirect('/step4');
  } catch (error) {
    console.error(error);
    res.status(500).send('Error updating room/lab');
  }
});

// Delete Room/Lab
router.post('/delete/:id', async (req, res) => {
  try {
    await Classroom.findByIdAndDelete(req.params.id);
    res.redirect('/step4');
  } catch (error) {
    console.error(error);
    res.status(500).send('Error deleting room/lab');
  }
});

// Navigation
router.post('/next', (req, res) => {
  res.redirect('/step5');
});

module.exports = router;
