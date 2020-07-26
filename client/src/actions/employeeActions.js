import {
  GET_EMPLOYEES,
  GET_REPORTING_TO_OPTION,
  ADD_EMPLOYEE,
  EMPLOYEE_EDITED,
  EMPLOYEE_DELETED,
  SELECT_EMPLOYEE,
  EMPLOYEES_LOADING,
  SERVICE_ERROR,
} from "./types";
import axios from "axios";

export const getEmployees = (id) => (dispatch) => {
  dispatch(setEmployeesLoading());
  axios
    .get("/api/employees")
    .then((resp) => {
      dispatch({
        type: GET_EMPLOYEES,
        payLoad: {
          employees: resp.data,
          employeeId: id ? id : resp.data[0]._id,
        },
      });
      dispatch({
        type: GET_REPORTING_TO_OPTION,
      });
      dispatch({
        type: SELECT_EMPLOYEE,
        payLoad: id,
      });
    })
    .catch((error) => {
      dispatch({ type: SERVICE_ERROR, payLoad: error });
    });
};
export const addEmployee = (employee) => (dispatch) => {
  dispatch(setEmployeesLoading());
  axios.post("/api/employees", employee).then((resp) => {
    dispatch({
      type: ADD_EMPLOYEE,
      payLoad: resp.data,
    });
  });
};
export const editEmployee = (id, employee) => (dispatch) => {
  dispatch(setEmployeesLoading());
  employee.id = id;
  axios.put(`/api/employees/${id}`, employee).then((resp) => {
    dispatch({
      type: EMPLOYEE_EDITED,
      payLoad: resp.data,
    });
  });
};
export const deleteEmployee = (id) => (dispatch) => {
  dispatch(setEmployeesLoading());
  axios.delete(`/api/employees/${id}`).then((resp) => {
    dispatch({
      type: EMPLOYEE_DELETED,
      payLoad: id,
    });
  });
};
export const downloadXLS = () => (dispatch) => {
  axios({
    url: "/api/employees/download",
    method: "GET",
    responseType: "blob", // important
  }).then((response) => {
    const url = window.URL.createObjectURL(new Blob([response.data]));
    const link = document.createElement("a");
    link.href = url;
    link.setAttribute("download", "employeeDetails.xlsx");
    document.body.appendChild(link);
    link.click();
  });
};
export const getReportingToOptions = () => (dispatch) => {
  return {
    type: GET_REPORTING_TO_OPTION,
  };
};
export const selectEmployee = (id) => {
  return {
    type: SELECT_EMPLOYEE,
    payLoad: id,
  };
};

export const setEmployeesLoading = () => {
  return {
    type: EMPLOYEES_LOADING,
    payLoad: true,
  };
};
