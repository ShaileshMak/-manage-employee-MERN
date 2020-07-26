import React, { Component } from 'react';
import { ListGroup, ListGroupItem} from 'reactstrap';


class EmployeeList extends Component {
    getEemployeeItem() {
        const employeesList = this.props.employees.map((employee) => {
            return (
                <ListGroupItem key={employee._id} 
                    tag="button" action 
                    onClick={this.props.onClick.bind(this, employee)}
                    style={{border:'none', fontWeight: employee._id === this.props.slectedId ? 'bold' : ''}}
                >{`${employee.firstName} ${employee.lastName}`}
                </ListGroupItem>
            )
        })

        return <ListGroup>{employeesList}</ListGroup>
    }

    render() {
        return (
            <div>
                {this.getEemployeeItem()}
            </div>
        )
    }
}


export default EmployeeList
