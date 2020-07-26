import React, { Component } from 'react'
import {
    Navbar,
    NavbarBrand,
    Nav,
    NavItem,
    NavbarToggler,
    Container,
    Collapse,
    NavLink
} from 'reactstrap';


import { connect } from 'react-redux';
import PropTypes from 'prop-types';

class Header extends Component {
    constructor(props) {
        super(props)

        this.state = {
            isOpen: false
        }
    }

    toggle = () => {
        this.setState({
            isOpen : !this.state.isOpen,
        })
    }

    showServiceError = () => this.props.employee.serviceError ? <dig>Service Error</dig> : '';

    render() {
        return (
            <div>
                <Navbar color="dark" dark expand="sm" className="mb-6">
                    <Container>
                        <NavbarBrand href="/">Employee Management</NavbarBrand>
                        <NavbarToggler onClick={this.toggle} />
                        <Collapse isOpen={this.state.isOpen} navbar>
                            <Nav className="ml-auto" navbar>
                                <NavItem>
                                    <NavLink href="/addEmployee">Add Employee</NavLink>
                                </NavItem>
                                <NavItem>
                                    <NavLink href="/about">About</NavLink>
                                </NavItem>
                                <NavItem>
                                    <NavLink href="/contactUs">Contact Us</NavLink>
                                </NavItem>
                            </Nav>
                        </Collapse>
                    </Container>
                </Navbar>
                {this.showServiceError()}
            </div>
        )
    }
}


Header.propTypes = {
    employee: PropTypes.object.isRequired
}

const mapStateToProps = (state) => ({
    employee: state.employee
});
export default connect(mapStateToProps)(Header)
