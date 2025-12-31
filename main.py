import itertools

game = [
    [0, 0, 0],
    [0, 0, 0],
    [0, 0, 0]
]


def game_board(game_map, player=0, row=0, column=0, just_display=False):
    try:
        if not just_display:
            game_map[row][column] = player

        print('   a  b  c')
        for count, row in enumerate(game_map):
            print(count, row)
        return game_map
    except IndexError as e:
        print('some error occurred:', e)
    except Exception as e:
        print('some other error occurred:', e)


def check_winner(game):
    # check for winner horizontally
    for row in game:
        if row.count(row[0]) == len(row) and row[0] != 0:
            print('Player 1 won horizontally! ')
            return True

    # Check for winner vertically
    for col in range(len(game)):
        check = []
        for row in game:
            check.append(row[col])
        if check.count(check[0]) == len(check) and check[0] != 0:
            print(f'Player {check[0]} won vertically!')
            return True

    # Check for winner diagonally
    diags = []
    for index, col in enumerate(range(len(game))):
        diags.append(game[index][col])

    if diags.count(diags[0]) == len(diags) and diags[0] != 0:
        print(f'Player {diags[0]} won diagonally! (\\)')
        return True

    diags = []
    for index, col in enumerate(reversed(range(len(game)))):
        diags.append(game[index][col])
    if diags.count(diags[0]) == len(diags) and diags[0] != 0:
        print(f'Player {diags[0]} won anti-diagonally! (/)')
        return True


play = True
while play:
    game = [[0, 0, 0],
            [0, 0, 0],
            [0, 0, 0]]
    game_won = False

    game = game_board(game, just_display=True)
    player_choice = itertools.cycle([1, 2])
    while not game_won:
        current_player = next(player_choice)
        print(f'Player {current_player} turn')
        column_choice = int(input('Choose a column: in range (0, 1, 2) '))
        row_choice = int(input('Choose a row: in range (0, 1, 2) '))
        game = game_board(game, current_player, row_choice, column_choice)

        if check_winner(game):
            game_won = True
            again = input('The game is over, would you like to play again? (y/n) ')
            if again.lower() == 'y':
                print('restarting')
            elif again.lower() == 'n':
                print('bye')
                play = False
            else:
                print('not a valid answer, bye')
                play = False
