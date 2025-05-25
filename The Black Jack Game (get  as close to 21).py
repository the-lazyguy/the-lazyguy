import random
def blackjet():
    user = list(map(int, input("enter you cards by space: ").split()))
    computer=[1,2,3,4,5,6,7,8,9,10]
    choic=random.choice(computer)
    print(choic)
    cont=input("Type 'y' to get another card,Type 'n' to pass:").lower()
    if cont == 'y':
        another=int(input("enter your another card :"))
        user.append(another)
        print(f"your final card is {user}")
        tuc=sum(user)
        choi=random.choice(computer)
        print(f"computer final hands{choic,choi}")
        tcc=choic+choi
        if tuc > 21:
            print("You lose.")
        elif tcc > 21:
            print( "You win")
        elif tcc > tuc:
            print("You lose.")
        elif tuc > tcc:
            print("You win!")
        else:
            print("It's a draw!")
    else:
        if cont == 'n':
            cho=random.choice(computer)
            print(f"computer final hands[{choic,cho}]")
            tcc=cho+choic
            tuc=sum(user)
            if tuc > 21:
                print("You lose.")
            elif tcc > 21:
                print("You win.")
            elif tcc > tuc:
                print("You lose.")
            elif tuc > tcc:
                print("You win!")
            else:
                print("It's a draw!")
    play_again = input("Do you want to play a game of Blackjack? Type 'y' or 'n': ").lower()
    if play_again == 'y':
        blackjet()
blackjet()