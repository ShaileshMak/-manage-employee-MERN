import React from 'react';
import { Route, BrowserRouter as Router} from 'react-router-dom';
import { Provider } from 'react-redux';

import './App.css';
import 'bootstrap/dist/css/bootstrap.min.css';
import store from './store';

//components
import Header from './components/header/Header';
import EmployeeContainer from './components/employeeContainer/EmployeeContainer';
import AddEditEmployee from './components/addEditEmployee/AddEditEmployee';
import About from './components/pages/About';
import ContactUs from './components/pages/ContactUs';

function App() {
  return (
    <Provider store={store}>
      <div className="App">
        <Header />
        <Router>
          <Route exact path="/" component={EmployeeContainer} />
          <Route exact path="/addEmployee" component={AddEditEmployee} />
          <Route exact path="/editEmployee/:employeeId" component={AddEditEmployee} />
          <Route exact path="/about" component={About} />
          <Route exact path="/contactUs" component={ContactUs} />
        </Router>
      </div>
    </Provider>
  );
}

export default App;
