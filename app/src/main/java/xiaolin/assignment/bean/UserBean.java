package xiaolin.assignment.bean;

import org.apache.poi.ss.usermodel.CellStyle;

public class UserBean {
    private String idcard;      //身份证
    private String name;        //名单
    private String sex;         //性别
    private String num;       //序号
    private int row;        //行
    private int add = 1;            //是否需要添加 0不需要添加 1 添加


    public UserBean() {
    }

    public UserBean(String name, String sex, String num, int row, String idcard) {
        this.name = name;
        this.sex = sex;
        this.num = num;
        this.row = row;
        this.idcard = idcard;
    }

    public int getAdd() {
        return add;
    }

    public void setAdd(int add) {
        this.add = add;
    }

    public String getIdcard() {
        return idcard;
    }

    public void setIdcard(String idcard) {
        this.idcard = idcard;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }


    public String getNum() {
        return num;
    }

    public void setNum(String num) {
        this.num = num;
    }

    public int getRow() {
        return row;
    }

    public void setRow(int row) {
        this.row = row;
    }

}
