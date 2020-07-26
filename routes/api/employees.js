const express = require("express");
const router = express.Router();
const Excel = require("exceljs");
const Employee = require("../../models/Employee");
const Moment = require("moment");
//@route GET api/employees
//@des Retrives employees
//@access Public
router.get("/", (req, resp) => {
  Employee.find()
    .sort({ firstName: -1 })
    .then((employees) => resp.json(employees))
    .catch((error) => console.log(eerror));
});

//@route POST api/employees
//@des Add employees
//@access Public
router.post("/", (req, resp) => {
  const employee = new Employee({
    firstName: req.body.firstName,
    lastName: req.body.lastName,
    dateOfJoining: req.body.dateOfJoining,
    department: req.body.department,
    reportingTo: req.body.reportingTo || "NA",
    skillSet: req.body.skillSet,
  });
  employee.save().then((employee) => resp.json(employee));
});

//@route UPDATE api/employees
//@des Update employees
//@access Public
router.put("/:id", (req, resp) => {
  Employee.findOneAndUpdate(
    { _id: req.params.id }, // Filter
    {
      $set: {
        firstName: req.body.firstName,
        lastName: req.body.lastName,
        dateOfJoining: req.body.dateOfJoining,
        department: req.body.department,
        reportingTo: req.body.reportingTo,
        skillSet: req.body.skillSet,
      },
    }
  )
    .then((employee) => {
      resp
        .status(200)
        .json({ message: "Update successful!", employee: employee });
    })
    .catch((err) => res.status(404).json({ success: false }));
});

// @route   DELETE api/employees/:id
// @desc    Delete A Item
// @access  Private
router.delete("/:id", (req, res) => {
  Employee.findOneAndDelete({ _id: req.params.id })
    .then((employee) =>
      employee.remove().then(() => res.json({ success: true }))
    )
    .catch((err) => res.status(404).json({ success: false }));
});

// @route   GET api/employees/download
// @desc    downloads xls
// @access  Private
router.get("/download", (req, res) => {
  Employee.find()
    .sort({ firstName: -1 })
    .then((employees) => {
      res.writeHead(200, {
        "Content-Disposition": 'attachment; filename="employeeDetails.xlsx"',
        "Transfer-Encoding": "chunked",
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      var workbook = new Excel.stream.xlsx.WorkbookWriter({ stream: res });
      var worksheet = workbook.addWorksheet("employees");
      workbook.creator = "Shailesh Mak";
      workbook.lastModifiedBy = "Me";
      workbook.created = new Date();
      worksheet.columns = [
        { header: "#", key: "no", width: 10, outlineLevel: 1 },
        { header: "Name", key: "name", width: 30, outlineLevel: 1 },
        { header: "Date of Joinging", key: "doj", width: 10, outlineLevel: 1 },
        { header: "Department", key: "department", width: 15, outlineLevel: 1 },
        {
          header: "Reporting To",
          key: "reportingTo",
          width: 30,
          outlineLevel: 1,
        },
        { header: "Skill Sets", key: "skillSets", width: 40, outlineLevel: 1 },
      ];
      employees.map((employee, index) => {
        worksheet.addRow({
          no: index + 1,
          name: `${employee.firstName} ${employee.lastName}`,
          doj: Moment(employee.dateOfJoining).format("YYYY-MM-DD"),
          department: employee.department,
          reportingTo: employee.reportingTo,
          skillSets: employee.skillSet,
        });
      });
      worksheet.commit();
      workbook.commit();
    })
    .catch((error) => console.log(eerror));
});

module.exports = router;
