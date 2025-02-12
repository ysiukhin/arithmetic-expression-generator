package org.home;

public enum ArithmeticOperation {
    ADD("+", 1) {
        @Override
        int calculate(int a, int b) {
            return a + b;
        }
    }
    , SUBTRACT("-", 1) {
        @Override
        int calculate(int a, int b) {
            return a - b;
        }
    }
//    , MULTIPLY("*", 0) {
//        @Override
//        int calculate(int a, int b) {
//            return a * b;
//        }
//    }
//    , DIVIDE("/") {
//        @Override
//        int calculate(int a, int b) {
//            return a / b;
//        }
//    }
    ;

    private final String operation;

    public int getOperationPriority() {
        return operationPriority;
    }

    private final int operationPriority;

    ArithmeticOperation(String operation, int operationPriority) {
        this.operation = operation;
        this.operationPriority = operationPriority;
    }

    @Override
    public String toString() {
        return " "+operation+" ";
    }

    abstract int calculate(int a, int b);
}
