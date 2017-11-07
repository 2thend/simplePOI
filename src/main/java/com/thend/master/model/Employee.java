package com.thend.master.model;

import java.io.Serializable;
import java.math.BigDecimal;


/**
 * 员工表
 */
public class Employee{
    /**
     * 员工ID
     */
	private Integer id;
    /**
     * 员工姓名
     */
	private String empName;
    /**
     * 员工编号
     */
	private String empOcde;
    /**
     * 手机号
     */
	private String mobile;
    /**
     * 住址
     */
	private String address;
    /**
     * 所属部门ID
     */
	private Integer departId;
    /**
     * 所属领导ID
     */
	private Integer leaderId;
    /**
     * 薪资
     */
	private BigDecimal salary;


	public Integer getId() {
		return id;
	}

	public void setId(Integer id) {
		this.id = id;
	}

	public String getEmpName() {
		return empName;
	}

	public void setEmpName(String empName) {
		this.empName = empName;
	}

	public String getEmpOcde() {
		return empOcde;
	}

	public void setEmpOcde(String empOcde) {
		this.empOcde = empOcde;
	}

	public String getMobile() {
		return mobile;
	}

	public void setMobile(String mobile) {
		this.mobile = mobile;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}

	public Integer getDepartId() {
		return departId;
	}

	public void setDepartId(Integer departId) {
		this.departId = departId;
	}

	public Integer getLeaderId() {
		return leaderId;
	}

	public void setLeaderId(Integer leaderId) {
		this.leaderId = leaderId;
	}

	public BigDecimal getSalary() {
		return salary;
	}

	public void setSalary(BigDecimal salary) {
		this.salary = salary;
	}

	public static final String ID = "id";

	public static final String EMP_NAME = "emp_name";

	public static final String EMP_OCDE = "emp_ocde";

	public static final String MOBILE = "mobile";

	public static final String ADDRESS = "address";

	public static final String DEPART_ID = "depart_id";

	public static final String LEADER_ID = "leader_id";

	public static final String SALARY = "salary";

	protected Serializable pkVal() {
		return this.id;
	}

}