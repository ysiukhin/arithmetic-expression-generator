package org.home;

public class ComplexOperand implements Operand {
    private ArithmeticExpression expression;

    public ComplexOperand(ArithmeticExpression expression) {
        this.expression = expression;
    }

    @Override
    public int getOperand() {
        return expression.getResult();
    }

    @Override
    public String toString() {
        return expression.getOp1().toString() + expression.getOperation().toString() + expression.getOp2().toString();
    }
}
