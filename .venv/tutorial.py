game = [
    [0, 0, 0],
    [0, 0, 0],
    [0, 0, 0]
]


def game_board(player=0, row=0, column=0, just_display=False):
    if not just_display:
        game[row][column] = player
    print('   a  b  c')
    for count, row in enumerate(game):
        print(count, row)


game_board(player=1, row=1, column=1)
