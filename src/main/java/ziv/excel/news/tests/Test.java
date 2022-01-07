package ziv.excel.news.tests;

import io.swagger.annotations.ApiModel;
import io.swagger.annotations.ApiModelProperty;

import java.util.Date;

@ApiModel
public  class Test {
    @ApiModelProperty("名称(长度在xx范围内)")
    private String name;
    @ApiModelProperty("年龄(0-25)")
    private int age;
    @ApiModelProperty("金钱(1-10)")
    private double money;
    @ApiModelProperty("生日(1999-2021)")
    private Date birthday;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public double getMoney() {
        return money;
    }

    public void setMoney(double money) {
        this.money = money;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    @Override
    public String toString() {
        return "Test{" +
                "name='" + name + '\'' +
                ", age=" + age +
                ", money=" + money +
                ", birthday=" + birthday +
                '}';
    }
}