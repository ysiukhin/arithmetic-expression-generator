package org.home;

public class SimpleOperand implements Operand {
    private Integer value;

    public SimpleOperand(Integer value) {
        this.value = value;
    }

    @Override
    public int getOperand() {
        return value;
    }

    @Override
    public String toString() {
        return "" + value;
    }
}
