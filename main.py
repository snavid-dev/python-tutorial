from multiprocessing.resource_sharer import stop

game = [
    [0, 0, 0],
    [0, 0, 0],
    [0, 0, 0]
]


def game_board(player=0, row=0, column=0, just_display=False):
    try:
        if not just_display:
            game[row][column] = player
        else:
            print('   a  b  c')
            for count, row in enumerate(game):
                print(count, row)
    except IndexError as e:
        print('some error occured:', e)




# Check for winner vertically
# def checkWinner(game):
#     for col in range(len(game)):
#         check = [];
#         for row in game:
#             check.append(row[col])
#         if check.count(check[0]) == len(check) and check[0] != 0:
#             print(f'Player {check[0]} won vertically!')
#             return


# Check for winner horizontally
# def checkWinner(game):
#     for row in game:
#         if row.count(row[0]) == len(row) and row[0] != 0:
#             print('Player 1 won horizontally!')
#             return


game_board(player=2, row=0, column=1)
game_board(player=2, row=1, column=1)
game_board(player=2, row=2, column=1)
game_board(just_display=True)
checkWinner(game)
