import React, { Component } from "react";
import Select from "react-select";
import { connect } from "react-redux";
import PropTypes from "prop-types";
import { Button } from "reactstrap";
import Moment from "moment";

import {
  getEmployees,
  getReportingToOptions,
  addEmployee,
  editEmployee,
} from "../../actions/employeeActions";
import { departmentOptions, departmentNames } from "../../utilities/constants";

class AddEditEmployee extends Component {
  UNSAFE_componentWillMount() {
    let curentFlow = "";
    if (
      this.props.match.path.indexOf("addEmployee") > -1 &&
      !this.props.employee.reportingToOptions.length > 0
    ) {
      this.props.getEmployees();
      curentFlow = "ADD";
    } else if (
      this.props.match.path.indexOf("editEmployee") > -1 &&
      this.props.match.params.employeeId
    ) {
      this.props.getEmployees(this.props.match.params.employeeId);
      curentFlow = "EDIT";
    }
    this.setState({
      flow: curentFlow,
    });
  }

  UNSAFE_componentWillReceiveProps(nextProps) {
    return !nextProps.employee.selectedEmployee
      ? null
      : this.props.match.path.indexOf("editEmployee") > -1
      ? this.setState({
          firstName: nextProps.employee.selectedEmployee.firstName,
          lastName: nextProps.employee.selectedEmployee.lastName,
          dateOfJoining: Moment(
            nextProps.employee.selectedEmployee.dateOfJoining
          ).format("YYYY-MM-DD"),
          departmentValue: nextProps.employee.selectedEmployee.department,
          reportToValue: nextProps.employee.selectedEmployee.reportingTo,
          skills: nextProps.employee.selectedEmployee.skillSet,
        })
      : "";
  }

  constructor(props) {
    super(props);
    this.state = defaultState;
  }

  onSubmit = () => {
    const employee = {
      firstName: this.state.firstName,
      lastName: this.state.lastName,
      dateOfJoining: this.state.dateOfJoining,
      department: this.state.departmentValue,
      reportingTo: this.state.reportToValue,
      skillSet: this.state.skills,
    };
    if (this.state.flow === "ADD") {
      this.props.addEmployee(employee);
    } else if (this.state.flow === "EDIT") {
      this.props.editEmployee(this.props.employee.selectedEmployeeId, employee);
    }
  };

  onChange = (event) => {
    this.setState({
      [event.target.name]: event.target.value,
    });
  };

  handleDepartmentChange = (selectedOption) => {
    this.setState({
      departmentValue: selectedOption.value,
    });
  };

  handleReportToChange = (selectedOption) => {
    this.setState({
      reportToValue: selectedOption.value,
    });
  };

  handleResetClick = (e) => {
    this.setState(defaultState);
  };

  render() {
    const flowText = this.state.flow === "ADD" ? "Add" : "Edit";
    return (
      <div id="addPayeeForm" className="container">
        <form onSubmit={this.onSubmit.bind(this)}>
          <div className="row">
            <div className="col-xs-12">
              <h2>{`${flowText} Employee details`}</h2>
            </div>
          </div>
          <div className="row form-field-container">
            <div className="col-xs-12 col-sm-4 col-lg-3 form-label-container">
              <label>First Name: </label>
            </div>
            <div className="col-xs-12 col-sm-8 col-lg-5 .offset-lg-4 form-input-container">
              <input
                id="firstNameInput"
                type="text"
                name="firstName"
                onChange={this.onChange}
                value={this.state.firstName}
              />
            </div>
          </div>
          <div className="row form-field-container">
            <div className="col-xs-12 col-sm-4 col-lg-3 form-label-container">
              <label>Last Name: </label>
            </div>
            <div className="col-xs-12 col-sm-8 col-lg-5 .offset-lg-4 form-input-container">
              <input
                type="text"
                name="lastName"
                onChange={this.onChange}
                value={this.state.lastName}
              />
            </div>
          </div>
          <div className="row form-field-container">
            <div className="col-xs-12 col-sm-4 col-lg-3 form-label-container">
              <label>Date of Joining: </label>
            </div>
            <div className="col-xs-12 col-sm-8 col-lg-5 .offset-lg-4 form-input-container">
              <input
                type="date"
                name="dateOfJoining"
                onChange={this.onChange}
                value={this.state.dateOfJoining}
              />
            </div>
          </div>
          <div className="row form-field-container">
            <div className="col-xs-12 col-sm-4 col-lg-3 form-label-container">
              <label>Department: </label>
            </div>
            <div className="col-xs-12 col-sm-8 col-lg-5 .offset-lg-4 form-input-container">
              <Select
                value={departmentOptions.filter(
                  ({ value }) => value === this.state.departmentValue
                )}
                onChange={this.handleDepartmentChange}
                options={departmentOptions}
              />
            </div>
          </div>
          <div className="row form-field-container">
            <div className="col-xs-12 col-sm-4 col-lg-3 form-label-container">
              <label>Report To: </label>
            </div>
            <div className="col-xs-12 col-sm-8 col-lg-5 .offset-lg-4 form-input-container">
              <Select
                value={this.props.employee.reportingToOptions.filter(
                  ({ value }) => value === this.state.reportToValue
                )}
                onChange={this.handleReportToChange}
                options={this.props.employee.reportingToOptions}
              />
            </div>
          </div>
          <div className="row form-field-container">
            <div className="col-xs-12 col-sm-4 col-lg-3 form-label-container">
              <label>Skills: </label>
            </div>
            <div className="col-xs-12 col-sm-8 col-lg-5 .offset-lg-4 form-input-container">
              <input
                type="text"
                name="skills"
                onChange={this.onChange}
                value={this.state.skills}
              />
            </div>
          </div>
          <div className="row form-field-container">
            <div className="col-xs-12 form-buttons-col">
              <Button
                color="secondary"
                size="lg"
                onClick={this.handleResetClick}
              >
                Reset
              </Button>
              <Button
                color="primary"
                size="lg"
                type="submit"
              >{`${flowText} Employee`}</Button>
            </div>
          </div>
        </form>
      </div>
    );
  }
}

const defaultState = {
  firstName: "",
  lastName: "",
  dateOfJoining: Moment(new Date()).format("YYYY-MM-DD"),
  departmentValue: departmentNames.DIGITAL,
  reportToValue: "",
  skills: "",
  flow: "",
};

AddEditEmployee.propTypes = {
  employee: PropTypes.object.isRequired,
  getEmployees: PropTypes.func.isRequired,
  addEmployee: PropTypes.func.isRequired,
  editEmployee: PropTypes.func.isRequired,
  getReportingToOptions: PropTypes.func.isRequired,
};

const mapStateToProps = (state) => ({
  employee: state.employee,
  getEmployees: state.getEmployees,
  addEmployee: state.getEmployees,
  editEmployee: state.editEmployee,
  getReportingToOptions: state.getReportingToOptions,
});

export default connect(mapStateToProps, {
  getEmployees,
  getReportingToOptions,
  addEmployee,
  editEmployee,
})(AddEditEmployee);
