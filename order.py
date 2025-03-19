class Order:
    def __init__(self, sku, name, cpf, order_number, qtd, phone):
        self.sku = str(sku)
        self.name = str(name)
        self.cpf = str(cpf).zfill(11)
        self.order_number = str(order_number)
        self.qtd = str(qtd)
        self.phone = str(phone)

    def __str__(self):
        return f"SKU: {self.sku}, Name: {self.name}, CPF: {self.cpf}, Order Number: {self.order_number}, Qtd: {self.qtd}, Phone: {self.phone}"
    
