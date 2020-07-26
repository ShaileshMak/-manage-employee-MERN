import React, { Component } from "react";
import Moment from "moment";
import PropTypes from "prop-types";
import { connect } from "react-redux";
import { Link } from "react-router-dom";

import { departmentOptions } from "../../utilities/constants";
import { deleteEmployee, downloadXLS } from "../../actions/employeeActions";

class EmployeeDetails extends Component {
  getDepartmentString() {
    let departmentString = departmentOptions.filter(
      (option) => this.props.selectedEmployee.department === option.value
    )[0].label;
    return departmentString;
  }

  deleteEmployee = (id) => {
    this.props.deleteEmployee(this.props.employee.selectedEmployeeId);
  };

  downloadXLS = () => {
    this.props.downloadXLS();
  };
  render() {
    const selectedEmployee = this.props.employee.selectedEmployee;
    return (
      <div>
        <div
          style={{ float: "right" }}
          className="col-xs-12 col-sm-4 offset-sm-8"
        >
          <Link onClick={this.downloadXLS}> Download XLS</Link>
          <br />
          <Link onClick={this.deleteEmployee}> Deleete</Link> |
          <Link to={`/editEmployee/${selectedEmployee._id}`}> Edit</Link>
        </div>
        <div className="detail-container">
          <span>
            <b>Name: </b>
          </span>
          <span>{`${selectedEmployee.firstName} ${selectedEmployee.lastName}`}</span>
        </div>
        <div className="detail-container">
          <span>
            <b>Department: </b>
          </span>
          <span>{this.getDepartmentString()}</span>
        </div>
        <div className="detail-container">
          <span>
            <b>Reporting to: </b>
          </span>
          <span>{selectedEmployee.reportingTo}</span>
        </div>
        <div className="detail-container">
          <span>
            <b>Date of Joinging: </b>
          </span>
          <span>
            {Moment(selectedEmployee.dateOfJoining).format("YYYY-MM-DD")}
          </span>
        </div>
        <div className="detail-container">
          <span>
            <b>Skills: </b>
          </span>
          <span>{selectedEmployee.skillSet}</span>
        </div>
      </div>
    );
  }
}

EmployeeDetails.propTypes = {
  employee: PropTypes.object.isRequired,
  deleteEmployee: PropTypes.func.isRequired,
  downloadXLS: PropTypes.func.isRequired,
};

const mapStateToProps = (state) => ({
  employee: state.employee,
  deleteEmployee: state.deleteEmployee,
  downloadXLS: state.downloadXLS,
});

export default connect(mapStateToProps, { deleteEmployee, downloadXLS })(
  EmployeeDetails
);
