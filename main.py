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


def checkWinner(game):
    for row in game:
        matched = True;
        for column in row:
            if column != row[0] or column == 0:
                matched = False
        if matched:
            print('Player 1 won!')
            return


game_board(player=1, row=1, column=1)
game_board(player=1, row=2, column=0)
game_board(player=1, row=2, column=2)
game_board(just_display=True)
checkWinner(game)
