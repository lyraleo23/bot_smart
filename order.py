class Order:
    def __init__(self, sku, name, cpf, order_number, qtd, phone):
        self.sku = sku
        self.name = name
        self.cpf = cpf
        self.order_number = order_number
        self.qtd = qtd
        self.phone = phone

    def __str__(self):
        return f"SKU: {self.sku}, Name: {self.name}, CPF: {self.cpf}, Order Number: {self.order_number}, Qtd: {self.qtd}, Phone: {self.phone}"
    
