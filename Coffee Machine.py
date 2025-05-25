water=300
Milk=200
coffee=100
Money=0
def report():
    global water, Milk, coffee, Money
    while True:
        user_input=input("what would you like ?(espresso/black/cappuccino)").lower()
        if user_input == "report":
            print("water:",water,"ml")
            print("cofee:",coffee,"g")
            print("Milk:",Milk,"ml")
            print("Money:",Money,"$")
        elif user_input == "black":
            if water<200 and coffee<24 and Milk<150:
                print("sorry water is not available")
                continue
            print("please insert a coins")
            penny = int(input("how many  penny?:"))
            quaters = int(input("how many  quaters?:"))
            nikcle = int(input("how many  nikcle?:"))
            dimes = int(input("how many  dimes?:"))
            Total = ((penny * 0.01) + (quaters * 0.25) + (nikcle * 0.05) + (dimes * 0.10))
            price = 2.5
            total = Total - price
            print(f"Here is change{total}")
            print("Here is black Enjoy")
            water-=200
            coffee-=24
            Milk-=150
            Money+=price
        elif user_input == "cappuccino":
            if water<250 and coffee<24 and Milk<100:
                print("sorry water/coffe/milk is not available")
                continue
            print("please insert a coins")
            penny = int(input("how many  penny?:"))
            quaters = int(input("how many  quaters?:"))
            nikcle = int(input("how many  nikcle?:"))
            dimes = int(input("how many  dimes?:"))
            Total = ((penny * 0.01) + (quaters * 0.25) + (nikcle * 0.05) + (dimes * 0.10))
            price = 4.6
            total = Total - price
            print(f"Here is change{total}")
            print("Here is cappuccino Enjoy")
            water-=250
            coffee-=24
            Milk-=150
            Money+=price

        elif user_input=="espresso":
            if water<200 and coffee<24 and Milk<100:
                print("sorry water is not available")
                continue
            print("please insert a coins")
            penny = int(input("how many  penny?:"))
            quaters = int(input("how many  quaters?:"))
            nikcle = int(input("how many  nikcle?:"))
            dimes = int(input("how many  dimes?:"))
            Total = ((penny * 0.01) + (quaters * 0.25) + (nikcle * 0.05) + (dimes * 0.10))
            price = 4.6
            total = Total - price
            print(f"Here is change{total}")
            print("Here is cappuccino Enjoy")
            water -= 250
            coffee -= 24
            Milk -= 150
            Money += price
        elif user_input == "uptade":
            add_extra_water=int(input("how many wwater do you add "))
            add_extra_coffee=int(input("how many coffee do you add "))
            add_extra_Milk=int(input("how many milk do you add "))
            water+=add_extra_water
            coffee+=add_extra_coffee
            Milk+=add_extra_Milk
report()