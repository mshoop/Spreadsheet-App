using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExpressionEngine
{
    public class ExpTree
    {
        public ExpTree() // Constructs expression tree from given expression
        {
            _variables = new Dictionary<string, double>(); 
            _expression = "0"; // default set to 0
            _root = Execute(_expression);
        }

        #region Node Classes
        private abstract class Node // Base node class
        {
            public Node Left, Right;
        }

        private class ConstantValueNode : Node // Represents numerical value
        {
            public double Value;
            public ConstantValueNode() { }
            public ConstantValueNode(double val) { Value = val; }
        }

        private class VariableNode : Node // Represents a variable in expression
        {
            public string VariableName;
            public VariableNode() { } 
            public VariableNode(string name) { VariableName = name; }
        }

        private class OperationNode : Node // Represents binary operator (+ - * /)
        {
            public char Operator;
            public OperationNode() { }
            public OperationNode(char node_operator) { Operator = node_operator; }
        }
        #endregion

        #region Fields
        private Node _root;
        private Dictionary<string, double> _variables;
        private string _expression;
        #endregion

        #region Properties
        public string ExpressionString 
        {
            get
            {
                return _expression;
            }
            set
            {
                _variables.Clear();
                _expression = value;
                _root = Execute(_expression);
            }
        }
        #endregion
        
        #region Methods
        public void SetVar(string varName, double varValue) // Sets variable in expression tree
        {
            _variables[varName] = varValue;
        }

        public List<string> GetVar() // Get variables in expression
        {
            return _variables.Keys.ToList();
        }

        public double Eval() // Evaluate expression
        {
            return EvalNode(_root);
        }

        private double EvalNode(Node node) // Evaluate each node in expression tree recursively
        {
            ConstantValueNode constantNode = node as ConstantValueNode;
            if (constantNode != null)
            {
                return constantNode.Value;
            }

            VariableNode variableNode = node as VariableNode;
            if (variableNode != null)
            {
                return _variables[variableNode.VariableName];
            }

            OperationNode operationNode = node as OperationNode; 
            if (operationNode != null)
            {
                switch (operationNode.Operator)
                {
                    case '+':
                        return EvalNode(operationNode.Left) + EvalNode(operationNode.Right);
                    case '-':
                        return EvalNode(operationNode.Left) - EvalNode(operationNode.Right);
                    case '*':
                        return EvalNode(operationNode.Left) * EvalNode(operationNode.Right);
                    case '/':
                        return EvalNode(operationNode.Left) / EvalNode(operationNode.Right);
                    default:
                        Console.WriteLine("Not a valid operator: " + operationNode.Operator);
                        return 0;
                }
            }
            Console.WriteLine("Not a valid node."); // error 
            return 0;
        }

        private Node Execute(string exp) // Turn string into an expression string to calculate
        {
            if (string.IsNullOrEmpty(exp)) // no operator, variables, constants, spaces, or () parenthesis brackets
            {
                return null;
            }
            if (exp[0] == '(') // if '(' at start
            {
                int countParenthesis = 0;
                for (int i = 0; i < exp.Length; i++) // find closing bracket
                {
                    if (exp[i] == '(')
                    {
                        countParenthesis++;
                    }
                    else if (exp[i] == ')') // if closing bracket found
                    {
                        countParenthesis--;
                        if (countParenthesis == 0) 
                        {
                            if (exp.Length - 1 != i)
                            {
                                break; 
                            }
                            else 
                            {
                                return Execute(exp.Substring(1, exp.Length - 2));
                            }
                        }
                    }
                }
            }
            // Execute each operator
            foreach (char op in new char[] { '+', '-', '*', '/' }) 
            {
                Node n = Execute(exp, op); // Overload Execute function
                if (n != null)
                {
                    return n;
                }
            }
            // Either constant or variable names at this point
            double value; 
            if (double.TryParse(exp, out value))
            {
                return new ConstantValueNode(value);
            }
            else
            {
                _variables[exp] = 0;
                return new VariableNode(exp);
            }
            //return null;
        }

        private Node Execute(string exp, char op) // Overloaded Execute function
        {
            int counter = 0;
            int i = exp.Length - 1; // start from end of string
            while (i >= 0)
            {
                if (exp[i] == '(')
                {
                    counter++;
                }
                else if (exp[i] == ')')
                {
                    counter--;
                }
                if (counter == 0 && exp[i] == op) // if hit correct operator
                {
                    // Form the subtree with the OperationNode being the root
                    OperationNode opNode = new OperationNode(op); 
                    opNode.Left = Execute(exp.Substring(0, i));
                    opNode.Right = Execute(exp.Substring(i + 1));
                    return opNode;
                }
                i--;
            }
            return null;
        }
        #endregion
    }
}
