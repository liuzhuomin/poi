package ziv.excel.news.imports;

public class PoiCell {
    private Class<?> javaType;
    private int cellType;
    private Object value;

    public Class<?> getJavaType() {
        return javaType;
    }

    public void setJavaType(Class<?> javaType) {
        this.javaType = javaType;
    }

    public int getCellType() {
        return cellType;
    }

    public void setCellType(int cellType) {
        this.cellType = cellType;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    public String getCellStrValue() {
        return this.value == null ? "" : this.value.toString();
    }

}