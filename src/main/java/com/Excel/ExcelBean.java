package com.Excel;


public class ExcelBean {
	private String number1;
	private String number2;
	private String operator;

	public String getNumber1() {
		return number1;
	}

	public void setNumber1(String number1) {
		this.number1 = number1;
	}

	public String getNumber2() {
		return number2;
	}

	public void setNumber2(String number2) {
		this.number2 = number2;
	}

	public String getOperator() {
		return operator;
	}

	public void setOperator(String operator) {
		this.operator = operator;
	}

	@Override
	public String toString() {
		return "ExcelBean [Number1=" + number1 + ", Number2=" + number2 + ", operator=" + operator + "]";
	}

}
