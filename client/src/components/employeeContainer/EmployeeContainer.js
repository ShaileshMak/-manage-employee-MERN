import React, { Component } from 'react';
import { connect } from 'react-redux';
import { getEmployees, selectEmployee } from '../../actions/employeeActions';
import PropTypes from 'prop-types';

import EmployeeList from '../employeeList/EmployeeList';
import EmployeeDetails from '../employeeDetails/EmployeeDetails';

class EmployeeContainer extends Component {
    onClick = (employee) => {
        this.props.selectEmployee(employee._id);
    }

    componentWillMount() {
        this.props.getEmployees();
    }

    render() {
        return (
            this.props.employee.selectedEmployee ? 
                <div className="container">
                    <div className="row">
                        <div className="col-xs-12 col-sm-5  col-lg-4">
                            <EmployeeList onClick={this.onClick} employees={this.props.employee.employees} slectedId={this.props.employee.selectedEmployeeId}/>
                        </div>
                        <div className="col-xs-12 col-sm-7 col-lg-8 employee-container">
                            <EmployeeDetails selectedEmployee={this.props.employee.selectedEmployee} />
                        </div>
                    </div>
                </div>
            : ''
        )
    }
}

EmployeeContainer.propTypes = {
    employee: PropTypes.object.isRequired,
    getEmployees: PropTypes.func.isRequired,
    selectEmployee: PropTypes.func.isRequired
}

const mapStateToProps = (state) => ({
    employee: state.employee
});

export default connect( mapStateToProps, { getEmployees, selectEmployee } )(EmployeeContainer)
