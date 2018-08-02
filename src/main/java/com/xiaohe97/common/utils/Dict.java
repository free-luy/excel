package com.xiaohe97.common.utils;

/**
 * 字典Entity
 * 
 * @author ThinkGem
 * @version 2013-05-15
 */
public class Dict {

	private String label; // 标签名
	private String value; // 数据值
	private String type; // 类型
	private String description;// 描述
	private Integer sort; // 排序

	public Dict() {
		super();
	}

	public String getLabel() {
		return label;
	}

	public void setLabel(String label) {
		this.label = label;
	}

	public String getValue() {
		return value;
	}

	public void setValue(String value) {
		this.value = value;
	}

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	public String getDescription() {
		return description;
	}

	public void setDescription(String description) {
		this.description = description;
	}

	public Integer getSort() {
		return sort;
	}

	public void setSort(Integer sort) {
		this.sort = sort;
	}

}