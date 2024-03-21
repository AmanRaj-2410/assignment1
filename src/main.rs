use std::collections::HashMap;
use std::fs::{self, File};
use std::io::{self, Write};
use calamine::{Reader, open_workbook, Xlsx};
use chrono::prelude::*;
use chrono::NaiveDate;
//use std::path::Path;

use clap::{App, Arg};
#[derive(Debug)]
struct Employee {
    emp_id: String,
    emp_name: String,
    dept_id: u32,
    mobile_no: String,
    email: String,
}

/*#[derive(Debug)]
struct Department {
    id: String,
    title: String,
    st: i32,
}

#[derive(Debug)]
struct Salary {
    emp_id: String,
    salid: String,
    saldate: String,
    sal:String,
    salstatus:String,
}

#[derive(Debug)]
struct Leave {
    emp_id: String,
    leaveid: String,
    leaveFrom: String,
    leaveTo:String,
    leaveType:String,
}*/

fn main() {
    let matches = App::new("Employee Reader")
        .arg(
            Arg::with_name("input")
                .short('i')
                .long("input")
                .value_name("INPUT_FILE")
                .help("Sets the input file to use")
                .takes_value(true)
                .required(true),
        )
        .arg(
            Arg::with_name("output")
                .short('o')
                .long("output")
                .value_name("OUTPUT_FILE")
                .help("Sets the output file to write")
                .takes_value(true)
                .required(true),
        )
        .get_matches();

        let path = "C:\\Users\\Hp\\Desktop\\Department.xlsx";
        
    

    // Extract file paths from command-line arguments
    let input_file_path = matches.value_of("input").unwrap();
    let output_file_path = matches.value_of("output").unwrap();

    let file_content = fs::read_to_string(input_file_path).expect("Unable to read input file");
    
    let mut employee_map: HashMap<u32, Employee> = HashMap::new();

    for line in file_content.lines().skip(1) {
        
        let fields: Vec<&str> = line.split('|').collect();

        
        let emp_id: u32 = fields[0].parse().expect("Invalid EmpId");
        let emp_name = fields[1].to_string();
        let dept_id = fields[2].parse().expect("Invalid DeptId");
        let mobile_no = fields[3].to_string();
        let email = fields[4].to_string();

        let employee = Employee {
            emp_id: emp_id.to_string().clone(),
            emp_name: emp_name.clone(),
            dept_id,
            mobile_no: mobile_no.clone(),
            email: email.clone(),
        };

        employee_map.insert(emp_id, employee);
    }

    // Open the Excel file
    let mut workbook: Xlsx<_> = open_workbook(path).unwrap();

    // HashMap to store department data
    let mut dep_data: HashMap<String, (String, String)> = HashMap::new();

    // Get the first worksheet in the workbook
    if let Ok(range) = workbook.worksheet_range("Sheet1") {
        // Skip the header row
        let mut rows = range.rows();
        rows.next(); // skip header

        // Iterate over the rows in the worksheet
        for row in rows {

            let dept_id = row[0].to_string();
            let dept_title = row[1].to_string();
            let dept_str = row[2].to_string();

            dep_data
                .entry(dept_id)
                //.and_modify(|d| d.)
                .or_insert((dept_title, dept_str));

        }
    } else {
        // Handle the error if the worksheet_range function returns an error
        println!("Error reading worksheet range");
        return;
    }

    let path1 = "C:\\Users\\Hp\\Desktop\\Salary.xlsx";

    let mut workbook: Xlsx<_> = open_workbook(path1).unwrap();

    // HashMap to store department data
    let mut sal_data: HashMap<String, (String, String, String, String)> = HashMap::new();

    // Get the first worksheet in the workbook
    if let Ok(range) = workbook.worksheet_range("Sheet1") {
        // Skip the header row
        let mut rows = range.rows();
        rows.next(); // skip header

        // Iterate over the rows in the worksheet
        for row in rows {
            let emp_id = row[0].to_string();
            let sal_id = row[1].to_string();
            let sal_date = row[2].to_string();
            let salary = row[3].to_string();
            let sal_status=row[4].to_string();

            sal_data
                .entry(emp_id)
                //.and_modify(|d| d.)
                .or_insert((sal_id,sal_date,salary,sal_status));
        }
    } else {
        // Handle the error if the worksheet_range function returns an error
        println!("Error reading worksheet range");
        return;
    }

    let path2 = "C:\\Users\\Hp\\Desktop\\Leave.xlsx";

    let mut workbook: Xlsx<_> = open_workbook(path2).unwrap();

    // HashMap to store department data
    let mut leave_data: HashMap<String, (String, String, String, String)> = HashMap::new();

    // Get the first worksheet in the workbook
    if let Ok(range) = workbook.worksheet_range("Sheet1") {
        // Skip the header row
        let mut rows = range.rows();
        rows.next(); // skip header

        // Iterate over the rows in the worksheet
        for row in rows {
            let emp_id = row[0].to_string();
            let leave_id = row[1].to_string();
            let leave_from = row[2].to_string();
            let leave_to = row[3].to_string();
            let leave_type=row[4].to_string();

            leave_data
                .entry(emp_id)
                //.and_modify(|d| d.)
                .or_insert((leave_id,leave_from,leave_to,leave_type));

        }
    } else {
        // Handle the error if the worksheet_range function returns an error
        println!("Error reading worksheet range");
        return;
    }

    // Write employee details to the output file
    write_employee_details(&employee_map,&dep_data,&sal_data,&leave_data, output_file_path).expect("Failed to write output file");
}

fn write_employee_details(employee_map: &HashMap<u32, Employee>,
    dep_data: &HashMap<String, (String, String)>, 
    sal_data: &HashMap<String, (String, String, String, String)>,
    leave_data: &HashMap<String, (String, String, String, String)>,
    output_file_path: &str) -> io::Result<()> {
    // Create the output file
    let mut output_file = File::create(output_file_path)?;

    // Write header to the file
    writeln!(
        output_file,
        "EmpId~#~EmpName~#~DeptTitle~#~MobileNo~#~Email~#~SalStatus~#~Leave"
    )?;

    println!("\n");

    let current_month = Local::now().month();
    let current_year = Local::now().year();

    // Write each employee's information to the file
    for (emp_id, employee) in employee_map {
        let dept_id_str = employee.dept_id.to_string();
        let (dept_title, _) = dep_data.get(&dept_id_str).unwrap();

        let sal_id_str = employee.emp_id.to_string();
        let (_, _, _, salstatus) = sal_data.get(&sal_id_str).unwrap();

        let leave_id_str = employee.emp_id.to_string();
        let (_, leave_from, _, _) = leave_data.get(&leave_id_str).unwrap();

        //println!("{:?}",leave_from);
        //println!("{:?}",leave_from);
        let leave_from_date = NaiveDate::parse_from_str(leave_from, "%d-%m-%Y").unwrap();
        let leave_day = leave_from_date.day();
        let leave_month = leave_from_date.month();
        let leave_year = leave_from_date.year();

        let curr_date = chrono::Utc::now();
        let year = curr_date.year();
        let month = curr_date.month();
        let day = curr_date.day();

        let mut cond:bool;

        //println!("{},{}",leave_year,year);
        if leave_month!=month{
            cond = false;
        }
        else{
            if leave_year!=year{
                cond = false;
            }
            else{
                cond = true;
            }
        }
        writeln!(
            output_file,
            "{}~#~{}~#~{}~#~{}~#~{}~#~{}~#~{}",
            emp_id, employee.emp_name, dept_title, employee.mobile_no, employee.email, salstatus, cond
        )?;
    }

    


    Ok(())

    
}
