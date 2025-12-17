package com.montran.model;

public class TestSuite {
	
	private String name;
	private boolean selected = true; // default
	
	public TestSuite(String name) {
		this.name = name;	
	}
	
	public String getName() {
		return name;
	}
	
	public boolean isSelected() {
		return selected;
	}
	
	public void setSelected(boolean selected) {
		this.selected = selected;
	}
	
	@Override
	public String toString() {
		return name;
	}
	
}
