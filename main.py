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
        if row[0] == row[1] == row[2] != 0:
            print(f'Player {row[0]} wins!')
            stop;



game_board(player=1, row=2, column=1)
game_board(player=1, row=2, column=0)
game_board(player=1, row=2, column=2)
game_board(just_display=True)
checkWinner(game)
