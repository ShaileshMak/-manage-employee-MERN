const mongoose= require('mongoose');
const Schema = mongoose.Schema;

const EmployeeSchema = new Schema({
    firstName: {
        type: String,
        required: true
    },
    lastName: {
        type: String,
        required: true
    },
    dateOfJoining: {
        type: Date,
        required: true
    },
    department: {
        type: String,
        required: true
    },
    reportingTo: {
        type: String,
        required: true
    },
    skillSet: {
        type: String,
        required: true
    }
});

module.exports = Employee = mongoose.model('employee', EmployeeSchema);