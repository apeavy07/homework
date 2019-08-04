--1. List the following details of each employee: employee number, last name, first name, gender, and salary.

SELECT emp.emp_no, emp.last_name, emp.first_name, emp.gender, sal.salary
FROM employees AS emp
	INNER JOIN salaries AS sal
	ON emp.emp_no = sal.emp_no;

--2. List employees who were hired in 1986.

SELECT emp_no, first_name, last_name, hire_date
FROM employees
WHERE EXTRACT(YEAR FROM hire_date) = '1986';

--3. List the manager of each department with the following information: department number, department name, the manager's employee number, last name, first name, and start and end employment dates.

SELECT d.dept_no, d.dept_name, emp.emp_no, emp.last_name, emp.first_name, emp.gender, dm.from_date, dm.to_date
FROM employees AS emp
	INNER JOIN dept_manager AS dm
	ON emp.emp_no = dm.emp_no
	INNER JOIN departments AS d
	ON dm.dept_no = d.dept_no;

--4. List the department of each employee with the following information: employee number, last name, first name, and department name.

SELECT emp.emp_no, emp.last_name, emp.first_name, d.dept_name
FROM employees AS emp
	INNER JOIN dept_emp AS de
	ON emp.emp_no = de.emp_no
	INNER JOIN departments AS d
	ON de.dept_no = d.dept_no;

--5. List all employees whose first name is "Hercules" and last names begin with "B."

SELECT *
FROM employees AS emp
WHERE first_name = 'Hercules'
AND last_name LIKE 'B%';

--6. List all employees in the Sales department, including their employee number, last name, first name, and department name.

SELECT emp.emp_no, emp.last_name, emp.first_name, d.dept_name
FROM employees AS emp

	INNER JOIN dept_emp AS de
	ON emp.emp_no = de.emp_no	
	INNER JOIN departments AS d
	ON de.dept_no = d.dept_no
	
WHERE d.dept_no = 'd007';

--7. List all employees in the Sales and Development departments, including their employee number, last name, first name, and department name.

SELECT emp.emp_no, emp.last_name, emp.first_name, d.dept_name
FROM employees AS emp

	INNER JOIN dept_emp AS de
	ON emp.emp_no = de.emp_no	
	INNER JOIN departments AS d
	ON de.dept_no = d.dept_no
	
WHERE d.dept_no IN ('d005', 'd007');

--8. In descending order, list the frequency count of employee last names, i.e., how many employees share each last name.

SELECT last_name, COUNT(last_name) AS last_name_count
FROM employees
GROUP BY last_name
ORDER BY COUNT(last_name) DESC;
