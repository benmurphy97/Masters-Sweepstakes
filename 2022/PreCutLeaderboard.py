import pandas as pd


url = 'https://www.espn.com/golf/leaderboard'
df = pd.read_html(io=url)[0]

df.drop(columns=['Unnamed: 0'], inplace=True) # drop first column

# replace WD with +99
df.loc[df['SCORE']=='WD', 'SCORE'] = '+99'

# replace E scores with +0
df.loc[df['SCORE']=='E', 'SCORE'] = '+0'

# remove the + sings from the overpar scores
df['SCORE'] = df['SCORE'].apply(lambda x: int(x[1:]) if x[0] != '-' else int(x))


# get picks from the excel and format to dict
# picks = pd.read_excel('Masters-Sweepstakes/2022/picks.xlsx', header=0)
# picks.set_index('Players', inplace=True)

# picks_d = {}
# for i in range(len(picks)):
#     player_name = str(picks.iloc[i].name)
#     golfers = list(picks.iloc[i].values)
#     picks_d[player_name] = golfers

picks = {
        'Eoghan Gleeson': ['Cameron Smith', 'Tyrrell Hatton', 'Adam Scott', 'Jason Kokrak', 'Billy Horschel', 'Lucas Glover'], 
        'Gary Gleeson': ['Justin Thomas', 'Sam Burns', 'Patrick Reed', 'Jason Kokrak', 'Ryan Palmer', 'Lucas Glover'], 
        'Niall Bobbett': ['Viktor Hovland', 'Shane Lowry', 'Justin Rose', 'Kevin Kisner', 'Billy Horschel', 'Hudson Swafford'], 
        'Eoin Higgins': ['Justin Thomas', 'Shane Lowry', 'Marc Leishman', 'Robert MacIntyre', 'Francesco Molinari', 'Charl Schwartzel'], 
        'Alan Holmes': ['Cameron Smith', 'Will Zalatoris', 'Bubba Watson', 'Jason Kokrak', 'Billy Horschel', 'Stewart Hagestad (a)'], 
        "Michael O'Houlihan": ['Scottie Scheffler', 'Bryson DeChambeau', 'Max Homa', 'Jason Kokrak', 'Min Woo Lee', 'Charl Schwartzel'], 
        'Conor Bobbett': ['Rory McIlroy', 'Louis Oosthuizen', 'Tiger Woods', 'Lee Westwood', 'Danny Willett', 'Charl Schwartzel'], 
        'Brian Bobbett': ['Cameron Smith', 'Shane Lowry', 'Tiger Woods', 'Webb Simpson', 'Billy Horschel', 'Charl Schwartzel'], 
        'David Bobbett': ['Justin Thomas', 'Louis Oosthuizen', 'Patrick Reed', 'Tom Hoge', 'Billy Horschel', 'Charl Schwartzel'], 
        'Paul Reynolds': ['Collin Morikawa', 'Sam Burns', 'Seamus Power', 'Kevin Kisner', 'Billy Horschel', 'Harry Higgs'], 
        'Eoin Donohoe': ['Cameron Smith', 'Corey Conners', 'Abraham Ancer', 'Lucas Herbert', 'Francesco Molinari', 'Harry Higgs'],
        'Shane Bobbett': ['Viktor Hovland', 'Louis Oosthuizen', 'Seamus Power', 'Lee Westwood', 'Billy Horschel', 'Charl Schwartzel'],
        'Dylan Walsh': ['Viktor Hovland', 'Shane Lowry', 'Abraham Ancer', 'Sepp Straka', 'Billy Horschel', 'James Piot (a)'], 
        'Tony Walsh': ['Rory McIlroy', 'Shane Lowry', 'Seamus Power', 'Lee Westwood', 'Billy Horschel', 'Fred Couples'], 
        'Ben Murphy': ['Cameron Smith', 'Shane Lowry', 'Marc Leishman', 'Robert MacIntyre', 'Billy Horschel', 'Hudson Swafford'], 
        'Gary O Reilly': ['Scottie Scheffler', 'Tyrrell Hatton', 'Tiger Woods', 'Lee Westwood', 'Zach Johnson', 'Charl Schwartzel'],
        'Brian Moran': ['Brooks Koepka', 'Sam Burns', 'Abraham Ancer', 'Tom Hoge', 'Padraig Harrington', 'Harry Higgs'], 
        'Sean Patchell': ['Brooks Koepka', 'Joaquin Niemann', 'Max Homa', 'Kevin Kisner', 'Francesco Molinari', 'Harry Higgs'],
        'Kev Moran ': ['Cameron Smith', 'Shane Lowry', 'Justin Rose', 'Christiaan Bezuidenhout', 'Padraig Harrington', 'Harry Higgs'], 
        "Lorcan O'Dea": ['Jon Rahm', 'Shane Lowry', 'Si Woo Kim', 'Lee Westwood', 'Padraig Harrington', 'Charl Schwartzel'], 
        'Fergal Moran': ['Cameron Smith', 'Tony Finau', 'Tiger Woods', 'Erik van Rooyen', 'Billy Horschel', 'Harry Higgs'], 
        'Shay Batelle': ['Scottie Scheffler', 'Joaquin Niemann', 'Tiger Woods', 'Kevin Kisner', 'Garrick Higgo', 'Bernhard Langer'], 
        'Barry Fitzpatrick': ['Brooks Koepka', 'Bryson DeChambeau', 'Adam Scott', 'Lee Westwood', 'Zach Johnson', 'Laird Shepherd (a)']
}

# correct spellings
# unique_player_picks = list(set([i for sublist in list(picks.values()) for i in sublist]))
# for i in unique_player_picks:
#     if len(df.loc[df['PLAYER']==i])==0:
#         print(i)
#         print(df.loc[df['PLAYER']==i])
#         print()


output_list = []
for i in picks.keys():
    output = {
        'Player': i,
        'Players making cut': len(df.loc[(df['PLAYER'].isin(picks[i])) & (df['SCORE'] <= 99)]),
        'Favourites': picks[i][0],
        'Favourites Score': df.loc[df['PLAYER'] == (picks[i][0])]['SCORE'].values[0],
        'Maybes': picks[i][1],
        'Maybes Score': df.loc[df['PLAYER'] == (picks[i][1])]['SCORE'].values[0],
        'Possibles': picks[i][2],
        'Possibles Score': df.loc[df['PLAYER'] == (picks[i][2])]['SCORE'].values[0],
        'Cut': picks[i][3],
        'Cut Score': df.loc[df['PLAYER'] == (picks[i][3])]['SCORE'].values[0],
        'Doubt It': picks[i][4],
        'Doubt It Score': df.loc[df['PLAYER'] == (picks[i][4])]['SCORE'].values[0],
        'Hail Mary': picks[i][5],
        'Hail Mary Score': df.loc[df['PLAYER'] == (picks[i][5])]['SCORE'].values[0],
        'Lowest 3 Scores': df.loc[df['PLAYER'].isin(picks[i])].iloc[:3]['SCORE'].sum()
    }
    output_list.append(output)
    
output_df = pd.DataFrame(output_list)

output_df['Rank'] = output_df['Lowest 3 Scores'].rank(method='min').astype(int)

sorted_leaderboard = output_df[['Rank',
                                'Player',
                                'Lowest 3 Scores',
                                'Players making cut',
                                'Favourites','Favourites Score', 
                                'Maybes', 'Maybes Score', 
                                'Possibles', 'Possibles Score',
                                'Cut', 'Cut Score',
                                'Doubt It', 'Doubt It Score',
                                'Hail Mary', 'Hail Mary Score'
                                ]].sort_values(by=['Lowest 3 Scores', 'Players making cut'], ascending=[True, False]).reset_index(drop=True)



# print(sorted_leaderboard)

# Write to excel file and perform formatting
# Create a Pandas Excel writer using XlsxWriter as the engine.
path_to_data_location = '/Users/benmurphy/OneDrive/Projects/Masters Sweepstakes/2022'

writer = pd.ExcelWriter(path_to_data_location + '/Masters Sweepstakes Leaderboard.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
sorted_leaderboard.to_excel(writer, sheet_name='Sheet1', na_rep="-", index=False)

# Get the xlsxwriter objects from the dataframe writer object.
workbook  = writer.book
worksheet = writer.sheets['Sheet1']

# centre the score columns
centre_cell_format = workbook.add_format()
centre_cell_format.set_align('center')
centre_cell_format.set_font_size(12)

worksheet.set_column(0, 0, 14, centre_cell_format)
worksheet.set_column(2, 2, 14, centre_cell_format)
for i in range(5, 16, 2):
    worksheet.set_column(i, i, 14, centre_cell_format)

centre_cell_format2 = workbook.add_format()
centre_cell_format2.set_align('center')
centre_cell_format2.set_font_size(12)
worksheet.set_column(3, 3, 18, centre_cell_format2)

# left align the player name columns
left_cell_format = workbook.add_format()
left_cell_format.set_align('left')
left_cell_format.set_font_size(14)

worksheet.set_column(1, 1, 16, left_cell_format)
for i in range(4, 15, 2):
    worksheet.set_column(i, i, 16, left_cell_format)
    

writer.save()