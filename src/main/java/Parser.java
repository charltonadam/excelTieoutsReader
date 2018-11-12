import org.apache.poi.sl.usermodel.Sheet;

public class Parser {

    public Node parseCell(String formula) {

        String currentVariable = "";
        Node root = null;
        boolean currentlyParsingVariable = false;

        for(int i = 0; i < formula.length(); i++) {
            char currentChar = formula.charAt(i);
            switch(currentChar) {
                case '+':
                    if(currentlyParsingVariable ) {
                        Node n = new Number(resolveVariable(currentVariable));
                        if(root == null) {
                            root = n;
                        } else {
                            root = root.addNode(n);
                        }
                        currentVariable = "";
                        currentlyParsingVariable = false;
                    }
                    root = root.addNode(new addition());
                    break;
                case '-':
                    if(currentlyParsingVariable ) {
                        Node n = new Number(resolveVariable(currentVariable));
                        if(root == null) {
                            root = n;
                        } else {
                            root = root.addNode(n);
                        }
                        currentVariable = "";
                        currentlyParsingVariable = false;
                    }
                    root = root.addNode(new subtraction());
                    break;
                case '*':
                    if(currentlyParsingVariable ) {
                        Node n = new Number(resolveVariable(currentVariable));
                        if(root == null) {
                            root = n;
                        } else {
                            root = root.addNode(n);
                        }
                        currentVariable = "";
                        currentlyParsingVariable = false;
                    }
                    root = root.addNode(new multiplication());
                    break;
                case '/':
                    if(currentlyParsingVariable ) {
                        Node n = new Number(resolveVariable(currentVariable));
                        if(root == null) {
                            root = n;
                        } else {
                            root = root.addNode(n);
                        }
                        currentVariable = "";
                        currentlyParsingVariable = false;
                    }
                    root = root.addNode(new division());
                    break;
                case '(':
                    if(currentlyParsingVariable ) {
                        Node n = new Number(resolveVariable(currentVariable));
                        if(root == null) {
                            root = n;
                        } else {
                            root = root.addNode(n);
                        }
                        currentVariable = "";
                        currentlyParsingVariable = false;
                    }

                    //this is where the fun begins
                    //we pass in another parser starting where the open paren is at, then skip to where that one closes.
                    //skip nested parens, those will be taken care of in further parsings.
                     root = root.addNode(parseCell(formula.substring(i + 1)));

                    int numberOfParens = 1;
                    while(numberOfParens > 0) {
                        i++;
                        if(formula.charAt(i) == ')') {
                            numberOfParens--;
                        } else if(formula.charAt(i) == '(') {
                            numberOfParens++;
                        }
                    }
                    break;
                case ')':
                    //if we ever hit this, then we are in a paren, and we need to leave.
                    return new parentheses(root);
                case ' ':
                    //spaces do nothing except mark where a variable stops.
                    if(currentlyParsingVariable ) {
                        Node n = new Number(resolveVariable(currentVariable));
                        if(root == null) {
                            root = n;
                        } else {
                            root = root.addNode(n);
                        }
                        currentVariable = "";
                        currentlyParsingVariable = false;
                    }
                    break;
                default:
                    currentVariable += currentChar;
                    currentlyParsingVariable = true;
                    break;
            }

        }

        if(currentlyParsingVariable ) {
            Node n = new Number(resolveVariable(currentVariable));
            if(root == null) {
                root = n;
            } else {
                root = root.addNode(n);
            }
        }


        return root;
    }


    private static double resolveVariable(String variable) {
        int row, column;
        if(variable.charAt(0) > 64) {
            //variable off of the excel file
            column = variable.charAt(0) - 65;
            row =  variable.charAt(1) - 49;
            return Main.testingSheet.getRow(row).getCell(column).getNumericCellValue();
        } else {
            return Double.parseDouble(variable);
        }
    }





    public abstract class Node {

        private Node parent;
        private int precedence;
        private Node leftChild;
        private Node rightChild;

        public void setPrecedence(int precedence) {
            this.precedence = precedence;
        }

        public Node getParent() {
            return parent;
        }

        public void setParent(Node parent) {
            this.parent = parent;
        }

        public Node getLeftChild() {
            return leftChild;
        }

        public void setLeftChild(Node leftChild) {
            this.leftChild = leftChild;
        }

        public Node getRightChild() {
            return rightChild;
        }

        public void setRightChild(Node rightChild) {
            this.rightChild = rightChild;
        }


        public int getPrecedence() {
            return precedence;
        };
        public abstract double execute();

        //returns the root node as far as they know
        public Node addNode(Node newNode) {
            if(precedence <= newNode.getPrecedence()) {
                //newNode should be higher than us on the tree
                newNode.setLeftChild(this);
                if(parent != null) {
                    newNode.setParent(parent);
                    parent = newNode;
                    return newNode;
                } else {
                    //base case, newNode is the root, return the new node
                    setParent(newNode);
                    return newNode;
                }
            } else {
                //newNode should be lower than us on the tree
                //theoretically, leftChild should always be filled with something at this point
                if(rightChild == null) {
                    rightChild = newNode;
                    newNode.setParent(this);
                } else {
                    //recursively go down to find its place
                    rightChild = rightChild.addNode(newNode);
                }
            }
            return this;
        }


    }

    private class Number extends Node {
        double value;

        public Number(double value) {
            this.value = value;
            setPrecedence(0);
        }



        public double execute(){
            return value;
        }

    }

    private class addition extends Node {

        public addition() {
            setPrecedence(3);
        }

        public double execute() {
            return getLeftChild().execute() + getRightChild().execute();
        }
    }

    private class subtraction extends Node {

        public subtraction() {
            setPrecedence(3);
        }

        public double execute() {
            return getLeftChild().execute() - getRightChild().execute();
        }
    }

    private class multiplication extends Node {

        public multiplication() {
            setPrecedence(2);
        }

        public double execute() {
            return getLeftChild().execute() * getRightChild().execute();
        }
    }

    private class division extends Node {

        public division() {
            setPrecedence(2);
        }

        public double execute() {
            return getLeftChild().execute() / getRightChild().execute();
        }
    }

    private class parentheses extends Node {

        public parentheses(Node n) {
            setLeftChild(n);
            setPrecedence(0);
        }



        public double execute() {
            return getLeftChild().execute();
        }
    }



}
