import {
  GET_EMPLOYEES,
  GET_REPORTING_TO_OPTION,
  SELECT_EMPLOYEE,
  EMPLOYEES_LOADING,
  EMPLOYEE_DELETED,
  DOWNLOAD_XLS,
  SERVICE_ERROR,
} from "../actions/types";

const initialState = {
  employees: [],
  reportingToOptions: [],
  loading: false,
  selectedEmployee: null,
  selectedEmployeeId: null,
  serviceError: false,
};

export default function (state = initialState, action) {
  switch (action.type) {
    case GET_EMPLOYEES:
      return {
        ...state,
        employees: action.payLoad.employees,
        loading: false,
        serviceError: false,
      };
    case GET_REPORTING_TO_OPTION:
      const newReportToOptions = state.employees.map((employee) => {
        return {
          label: `${employee.firstName} ${employee.lastName}`,
          value: `${employee.firstName} ${employee.lastName}`,
        };
      });
      return {
        ...state,
        reportingToOptions: newReportToOptions,
      };
    case SELECT_EMPLOYEE:
      const id =
        action.payLoad || (state.employees && state.employees[0]._id) || null;
      return {
        ...state,
        selectedEmployeeId: id,
        selectedEmployee:
          (state.employees &&
            state.employees.filter((employee) => employee._id === id)[0]) ||
          null,
      };
    case EMPLOYEE_DELETED:
      const employees =
        state.employees &&
        state.employees.filter((employee) => employee._id !== action.payLoad);
      const selectedEmployeeId =
        (employees && employees[0] && employees[0]._id) || null;
      const selectedEmployee =
        (employees && employees[0] && employees[0]) || null;
      return {
        ...state,
        employees: employees,
        selectedEmployeeId: selectedEmployeeId,
        selectedEmployee: selectedEmployee,
      };
    case DOWNLOAD_XLS:
      return {
        ...state,
      };
    case EMPLOYEES_LOADING:
      return {
        ...state,
        loading: action.payLoad,
      };
    case SERVICE_ERROR:
      return {
        ...state,
        loading: false,
        serviceError: true,
      };
    default:
      return state;
  }
}
