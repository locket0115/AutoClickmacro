class Order:
    def __init__(self, bs: bool, price: float, method: str, quantity: int):
        self.bs = bs # 0: 매도, 1: 매수
        self.price = price
        self.method = method
        self.quantity = quantity

    def __repr__(self):
        return f"bs: {'매수' if self.bs else '매도'}, price: {self.price}, method: {self.method}, quantity: {self.quantity}"

class Stocks:
    def __init__(self, stock: str, orders: list[Order]):
        self.stock = stock
        self.orders = orders

    def add_order(self, order: Order):
        self.orders.append(order)

    def order_len(self):
        return len(self.orders)

    def __repr__(self):
        return f"stock: {self.stock}, orders: {self.orders}"
    

class Commands:
    def __init__(self, name: str):
        self.name = name
        self.command = []

    def add_command(self, command):
        self.command.append(command)