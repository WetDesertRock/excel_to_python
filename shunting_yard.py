import ast
from util import *

from openpyxl.formula.tokenizer import Token, Tokenizer


class ShuntingYardException(Exception):
    pass

# TODO: These token classes really should not exist in their current form.
# There should just be 1 token class, and then some sort of way to easily
# create it based on the symbol.


class OperatorToken:
    precedence: int = -1
    associativity: int = 0  # 0 is left associativity, 1 is right


class BinaryOperatorToken(OperatorToken):
    symbol = None
    ast_fun = None


class UnaryOperatorToken(OperatorToken):
    symbol = None
    ast_fun = None
    associativity = 1
    precedence = 10


class CompareOperatorToken(OperatorToken):
    symbol = None
    ast_fun = None


class UnarySubOperatorToken(UnaryOperatorToken):
    symbol = "-u"
    ast_fun = ast.USub


class UnaryAddOperatorToken(UnaryOperatorToken):
    symbol = "+u"
    ast_fun = ast.UAdd


class AddOperatorToken(BinaryOperatorToken):
    symbol = "+"
    precedence = 3
    ast_fun = ast.Add


class SubOperatorToken(BinaryOperatorToken):
    symbol = "-"
    precedence = 3
    ast_fun = ast.Sub


class MultOperatorToken(BinaryOperatorToken):
    symbol = "*"
    precedence = 5
    ast_fun = ast.Mult


class DivOperatorToken(BinaryOperatorToken):
    symbol = "/"
    precedence = 5
    ast_fun = ast.Div


class PowOperatorToken(BinaryOperatorToken):
    symbol = "^"
    precedence = 7
    ast_fun = ast.Pow
    associativity = 1


class LtOperatorToken(CompareOperatorToken):
    symbol = "<"
    precedence = 0
    ast_fun = ast.Lt


class LtEqOperatorToken(CompareOperatorToken):
    symbol = "<="
    precedence = 0
    ast_fun = ast.LtE


class GtOperatorToken(CompareOperatorToken):
    symbol = ">"
    precedence = 0
    ast_fun = ast.Gt


class GtEqOperatorToken(CompareOperatorToken):
    symbol = "<="
    precedence = 0
    ast_fun = ast.GtE


class EqOperatorToken(CompareOperatorToken):
    symbol = "="
    precedence = 0
    ast_fun = ast.Eq


class NotEqOperatorToken(CompareOperatorToken):
    symbol = "<>"
    precedence = 0
    ast_fun = ast.NotEq


class ParenOperatorToken(OperatorToken):
    pass


class FuncOperatorToken(OperatorToken):
    def __init__(self, name):
        self.name = name


class SeperatorOperatorToken(OperatorToken):
    pass


operators = [
    AddOperatorToken, SubOperatorToken, MultOperatorToken, DivOperatorToken, PowOperatorToken,
    LtOperatorToken, LtEqOperatorToken, GtOperatorToken, GtEqOperatorToken, EqOperatorToken, NotEqOperatorToken,
    UnarySubOperatorToken, UnaryAddOperatorToken
]
operator_symbol_map = {op.symbol: op for op in operators}

# TODO: redo all of this nasty.


class ShuntingYard():
    def __init__(self, tokens, variable_map):
        self.output = []
        self.ops = []
        self.tokens = tokens
        self.variable_map = variable_map

        self.token_idx = 0

    def apply_op(self, op):
        assert op.ast_fun != None, f"Attempting to add invalid operator: {op}"

        if isinstance(op, CompareOperatorToken):
            right = self.output.pop()
            left = self.output.pop()
            compare_op = ast.Compare(
                left=left,
                ops=[op.ast_fun()],
                comparators=[right]
            )
            self.output.append(compare_op)
        elif isinstance(op, UnaryOperatorToken):
            operand = self.output.pop()
            self.output.append(ast.UnaryOp(
                op=op.ast_fun(),
                operand=operand
            ))
        elif isinstance(op, BinaryOperatorToken):
            right = self.output.pop()
            left = self.output.pop()
            bin_op = ast.BinOp(left=left, right=right, op=op.ast_fun())
            self.output.append(bin_op)
        else:
            raise ShuntingYardException(f"Unable to apply operator {op}, unknown type.")

    def add_op(self, op):
        while len(self.ops) > 0 and (
            self.ops[-1].precedence > op.precedence or (op.associativity ==
                                                        0 and op.precedence == self.ops[-1].precedence)
        ):
            self.apply_op(self.ops.pop())

        self.ops.append(op)

    def resolve_variable_lookup(self, cell_coordinate):
        cell_info = self.variable_map.get(cell_coordinate)
        if cell_info == None:
            raise ShuntingYardException(f"Unable to find CellInfo for {cell_coordinate}")

        var_ast = ast.Attribute(value=ast.Name(id="self", ctx=ast.Load()),
                                attr=cell_info.variable_name, ctx=ast.Load())
        if cell_info.is_formula():
            var_ast = ast.Call(func=var_ast, args=[], keywords=[])

        return var_ast

    def process_token_operand(self, token):
        if token.subtype == Token.TEXT:
            self.output.append(ast.Constant(token.value))
        elif token.subtype == Token.NUMBER:
            self.output.append(ast.Constant(float(token.value)))
        elif token.subtype == Token.RANGE:
            cells = [c for c in excel_range_iter(token.value)]
            if len(cells) > 0:
                for cell in cells:
                    var_ast = self.resolve_variable_lookup(cell)

                    self.output.append(var_ast)
                    self.ops.append(SeperatorOperatorToken())

                self.ops.pop()  # Remove the excess seperator

    def process_token_infix(self, token):
        bin_op = operator_symbol_map[token.value]()
        self.add_op(bin_op)

    def process_token_paren(self, token):
        if token.subtype == Token.OPEN:
            self.ops.append(ParenOperatorToken())
        else:
            while len(self.ops) > 0 and not isinstance(self.ops[-1], ParenOperatorToken):
                self.apply_op(self.ops.pop())

            assert len(self.ops) > 0, "Unable to find left paren"
            self.ops.pop()  # Throw away our '('

    def process_token_func(self, token):
        if token.subtype == Token.OPEN:
            self.ops.append(FuncOperatorToken(token.value.rstrip("(")))

            lookahead = self.look_ahead()
            if lookahead.type == Token.FUNC and lookahead.subtype == Token.CLOSE:
                op = self.ops.pop()
                fun = ast.Call(
                    func=ast.Name(id=op.name, ctx=ast.Load()),
                    args=[],
                    keywords=[]
                )
                self.output.append(fun)
                self.token_idx += 1

        elif token.subtype == Token.CLOSE:
            arg_list = []
            while len(self.ops) > 0 and not isinstance(self.ops[-1], FuncOperatorToken):
                op = self.ops.pop()
                if isinstance(op, SeperatorOperatorToken):
                    arg_list.append(self.output.pop())
                else:
                    self.apply_op(op)

            arg_list.append(self.output.pop())
            arg_list.reverse()
            op = self.ops.pop()
            fun = ast.Call(
                func=ast.Name(id=op.name, ctx=ast.Load()),
                args=arg_list,
                keywords=[]
            )
            self.output.append(fun)

    def process_token_prefix(self, token):
        op = operator_symbol_map[token.value + "u"]()
        self.add_op(op)

    def process_token(self, token):
        if token.type == Token.LITERAL:
            self.output.append(ast.Constant(token.value))
        elif token.type == Token.OPERAND:
            self.process_token_operand(token)
        elif token.type == Token.OP_IN:
            self.process_token_infix(token)
        elif token.type == Token.PAREN:
            self.process_token_paren(token)
        elif token.type == Token.FUNC:
            self.process_token_func(token)
        elif token.type == Token.SEP:
            self.ops.append(SeperatorOperatorToken())
        elif token.type == Token.WSPACE:
            pass
        elif token.type == Token.OP_PRE:
            self.process_token_prefix(token)

        else:
            raise ShuntingYardException(
                f"Unexpected token type. Token: {token.type}, subype: {token.subtype}, value: {token.value}")

    def look_ahead(self):
        return self.tokens[self.token_idx + 1]

    def process(self):
        while self.token_idx < len(self.tokens):
            token = self.tokens[self.token_idx]
            self.process_token(token)
            self.token_idx += 1

        while len(self.ops):
            self.apply_op(self.ops.pop())

        assert len(self.output) == 1, "Output expr list does not have exactly one member"

        return self.output[0]


class FunctionRewriter(ast.NodeTransformer):
    rewrite_const_map = {
        "PI": ("math", "pi")
    }
    rewrite_func_map = {
        "SIN": ("math", "sin"),
        "COS": ("math", "cos"),
        "ACOS": ("math", "acos"),
        "EXP": ("math", "exp"),
    }

    def visit_Call(self, node):
        super().generic_visit(node)

        # If this isn't a standard name lookup, assume it has been normalized already
        if isinstance(node.func, ast.Attribute):
            return node

        func_id = node.func.id

        if func_id in FunctionRewriter.rewrite_const_map:
            module, const = FunctionRewriter.rewrite_const_map[func_id]
            return ast.Attribute(
                value=ast.Name(id=module, ctx=ast.Load()),
                attr=const,
                ctx=ast.Load()
            )

        if func_id in FunctionRewriter.rewrite_func_map:
            module, func_name = FunctionRewriter.rewrite_func_map[func_id]
            node.func = ast.Attribute(
                value=ast.Name(id=module, ctx=ast.Load()),
                attr=func_name,
                ctx=ast.Load()
            )
            return node

        if func_id == "IF":
            test = node.args[0]
            body = node.args[1]
            orelse = node.args[2]

            return ast.If(
                test=test,
                body=[ast.Return(value=body)],
                orelse=[ast.Return(value=orelse)],
                # lineno=node.lineno
            )

        return node


if __name__ == "__main__":
    code = "=-(1+1)"
    tokens = Tokenizer(code)
    shunting_yard = ShuntingYard(tokens.items, {})
    # print("\t".join([f"{t.type}-{t.subtype}-{t.value}" for t in tokens.items]))
    node = ast.fix_missing_locations(FunctionRewriter().visit(shunting_yard.process()))
    # print(node)
    print(ast.dump(node, indent=2, include_attributes=True))
    print(">>>", ast.unparse(node))
