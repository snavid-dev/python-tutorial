game = [
    [0, 0, 0],
    [0, 0, 0],
    [0, 0, 0]
]


def game_board(player=0, row=0, column=0, just_display=False):
    try:
        if not just_display:
            game[row][column] = player
            print('   a  b  c')
            for count, row in enumerate(game):
                print(count, row)
    except IndexError as e:
        print('some error occured:', e)


game_board(player=1, row=5, column=1)