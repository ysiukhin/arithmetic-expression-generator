package org.home;

public class ArithmeticExpression {
    private ArithmeticOperation operation;
    private Operand op1;
    private Operand op2;


    public ArithmeticOperation getOperation() {
        return operation;
    }

    public Operand getOp1() {
        return op1;
    }

    public Operand getOp2() {
        return op2;
    }

    public int getResult() {
        return result;
    }


    private int result;

    public ArithmeticExpression(ArithmeticOperation op, Operand i, Operand j, int result) {
        this.operation = op;
        this.op1 = i;
        this.op2 = j;
        this.result = result;
    }

    @Override
    public String toString() {
        return op1 + operation.toString() + op2 + " = " + result;
    }
}
