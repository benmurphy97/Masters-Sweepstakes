import pandas as pd
import ssl
import xlsxwriter 
ssl._create_default_https_context = ssl._create_unverified_context

url = 'https://www.augusta.com/masters/leaderboard'

df = pd.read_html(io=url)[0]

df = df.iloc[2:].reset_index(drop=True) # drop first 2 not needed rows

new_header = df.iloc[0] #grab the first row for the header
df = df[1:] #take the data less the header row
df.columns = new_header #set the header row as the df header

# rename duplicate column
df.columns.values[2] = "LATEST_SCORE"

# replace E scores with +0
df.loc[df['LATEST_SCORE']=='E', 'LATEST_SCORE'] = '+0'

# remove the + sings from the overpar scores
df['LATEST_SCORE'] = df['LATEST_SCORE'].apply(lambda x: int(x[1:]) if x[0] != '-' else int(x))
# df['LATEST_SCORE'] = df['LATEST_SCORE'].astype(int)
df.head(10)

picks = {
    'Jimmy ODowd': ['Jordan Spieth', 'Tommy Fleetwood', 'Ian Poulter', 'Danny Willett', 'Martin Laird', 'Tyler Strafaci'],
    'Ben Murphy': ['Bryson DeChambeau', 'Daniel Berger', 'Matt Wallace', 'Dylan Frittelli', 'Bernd Wiesberger', 'Mike Weir'],
    'Ciarán Cahill': ['Justin Thomas', 'Sungjae Im', 'Abraham Ancer', 'Cameron Champ', 'Jimmy Walker', 'Tyler Strafaci'],
    'Paddy ODowd': ['Justin Thomas', 'Jason Day', 'Abraham Ancer', 'Cameron Champ', 'Jimmy Walker', 'Tyler Strafaci'],
    'Eoghan Gleeson': ['Justin Thomas', 'Sergio Garcia', 'Billy Horschel', 'Phil Mickelson', 'Bernd Wiesberger', 'Mike Weir'],
    'Jacob Henry-Hayes': ['Patrick Cantlay', 'Matthew Fitzpatrick', 'Billy Horschel', 'Phil Mickelson', 'C.T. Pan', 'Tyler Strafaci'],
    'Eoin Higgins': ['Justin Thomas', 'Tommy Fleetwood', 'Justin Rose', 'Cameron Champ', 'Henrik Stenson', 'Vijay Singh'],
    'Matthew Odlum': ['Dustin Johnson', 'Webb Simpson', 'Abraham Ancer', 'Sebastian Munoz', 'Bernhard Langer', 'Vijay Singh'],
    'Sean Patchell': ['Dustin Johnson', 'Adam Scott', 'Abraham Ancer', 'Kevin Na', 'Charl Schwartzel', 'Jose Maria Olazabal'],
    'Barry Fitzpatrick': ['Jordan Spieth', 'Tommy Fleetwood', 'Gary Woodland', 'Jimmy Walker', 'Bernhard Langer', 'Jim Herman'],
    'Brian Bobbett': ['Justin Thomas', 'Matthew Fitzpatrick', 'Francesco Molinari', 'Ryan Palmer', 'Charl Schwartzel', 'Robert Streb'],
    'Alan Holmes': ['Justin Thomas', 'Daniel Berger', 'Billy Horschel', 'Gary Woodland', 'Martin Laird', 'Vijay Singh'],
    'Rory Collins': ['Jordan Spieth', 'Daniel Berger', 'Brian Harman', 'Robert MacIntyre', 'Charl Schwartzel', 'Vijay Singh'],
    'Alex Marshall': ['Justin Thomas', 'Webb Simpson', 'Brian Harman', 'Sebastian Munoz', 'C.T. Pan', 'Charles Osborne'],
    'Ryan Johnson': ['Dustin Johnson', 'Sergio Garcia', 'Justin Rose', 'Phil Mickelson', 'Charl Schwartzel', 'Vijay Singh'],
    'James Browne': ['Jordan Spieth', 'Matthew Fitzpatrick', 'Matt Wallace', 'Phil Mickelson', 'Jimmy Walker', 'Jose Maria Olazabal'],
    'Faolan Crowe': ['Jon Rahm', 'Webb Simpson', 'Matthew Wolff', 'Phil Mickelson', 'C.T. Pan', 'Mike Weir'],
    # 'Kevin Moran': ['Xander Schauffele', 'Sergio Garcia', 'Matt Kuchar', 'Christiaan Bezuidenhout', 'C.T. Pan', 'Tyler Strafaci'],
    'Brian Moran': ['Collin Morikawa', 'Tyrrell Hatton', 'Si Woo Kim', 'Gary Woodland', 'C.T. Pan', 'Tyler Strafaci'],
    'Fergal Moran': ['Jordan Spieth', 'Scottie Scheffler', 'Matt Kuchar', 'Dylan Frittelli', 'Martin Laird', 'Jose Maria Olazabal'],
    'Jack Dignam': ['Brooks Koepka', 'Shane Lowry', 'Billy Horschel', 'Zach Johnson', 'Henrik Stenson', 'Jose Maria Olazabal'],
    # 'Ted ODonoghue': ['Lee Westwood', 'Corey Conners', 'Matthew Wolff', 'Phil Mickelson', 'C.T. Pan', 'Robert Streb'],
    'Jack Byrne': ['Jon Rahm', 'Adam Scott', 'Billy Horschel', 'Sebastian Munoz', 'Charl Schwartzel', 'Jose Maria Olazabal'],
    # 'Jamie Murphy': ['Jordan Spieth', 'Sungjae Im', 'Billy Horschel', 'Kevin Kisner', 'Fred Couples', 'Mike Weir'],
    'Dave Bobbett': ['Bryson DeChambeau', 'Daniel Berger', 'Billy Horschel', 'Gary Woodland', 'Charl Schwartzel', 'Vijay Singh'],
    'Dylan Walsh': ['Justin Thomas', 'Shane Lowry', 'Marc Leishman', 'Phil Mickelson', 'Brendon Todd', 'Vijay Singh'],
    'Alan Horgan': ['Jordan Spieth', 'Daniel Berger', 'Si Woo Kim', 'Dylan Frittelli', 'Bernd Wiesberger', 'Jose Maria Olazabal'],
    'Fiachra Keane': ['Justin Thomas', 'Corey Conners', 'Francesco Molinari', 'Phil Mickelson', 'Charl Schwartzel', 'Sandy Lyle'],
    # 'Mark Kirwan': ['Dustin Johnson', 'Sungjae Im', 'Abraham Ancer', 'Phil Mickelson', 'Jimmy Walker', 'Tyler Strafaci'],
    'Niall Bobbett': ['Jordan Spieth', 'Daniel Berger', 'Matt Kuchar', 'Ryan Palmer', 'Bernd Wiesberger', 'Robert Streb'],
    'Shay Moran': ['Jordan Spieth', 'Sergio Garcia', 'Brian Harman', 'Phil Mickelson', 'Martin Laird', 'Vijay Singh']
}


output_list = []
for i in picks.keys():
    output = {
        'Player': i,
        
        'Players making cut': len(df.loc[(df['NAME'].isin(picks[i])) & (df['POS'] != 'CUT')]),
        
        'Favourites': picks[i][0],
#         'Favourites Score': df.loc[df['NAME'] == (picks[i][0])]['LATEST_SCORE'].values[0],
        
        'Maybes': picks[i][1],
#         'Maybes Score': df.loc[df['NAME'] == (picks[i][1])]['LATEST_SCORE'].values[0],
        
        'Possibles': picks[i][2],
#         'Possibles Score': df.loc[df['NAME'] == (picks[i][2])]['LATEST_SCORE'].values[0],
        
        'Cut': picks[i][3],
        'Cut Score': df.loc[df['NAME'] == (picks[i][3])]['LATEST_SCORE'].values[0],
        
        'Doubt It': picks[i][4],
        'Doubt It Score': df.loc[df['NAME'] == (picks[i][4])]['LATEST_SCORE'].values[0],
        
        'Past It': picks[i][5],
        'Past It Score': df.loc[df['NAME'] == (picks[i][5])]['LATEST_SCORE'].values[0],
        
        'Lowest 3 Scores': df.loc[(df['NAME'].isin(picks[i])) & (df['POS'] != 'CUT')].iloc[:3]['LATEST_SCORE'].sum()
    }
    
    # get favourite score if they made the cut
    favourites_index = 0
    favourites_score = df.loc[df['NAME'] == (picks[i][favourites_index])]['LATEST_SCORE'].values[0]
    position = df.loc[df['NAME'] == (picks[i][favourites_index])]['POS'].values[0]
    if position == 'CUT':
        output['Favourites Score'] = 'CUT'
    elif position == 'DQ':
        output['Favourites Score'] = 'DQ'
    else:
        output['Favourites Score'] = favourites_score
        
    # get maybes score if they made the cut
    maybes_index = 1
    maybes_score = df.loc[df['NAME'] == (picks[i][maybes_index])]['LATEST_SCORE'].values[0]
    position = df.loc[df['NAME'] == (picks[i][maybes_index])]['POS'].values[0]
    if position == 'CUT':
        output['Maybes Score'] = 'CUT'
    elif position == 'DQ':
        output['Maybes Score'] = 'DQ'
    else:
        output['Maybes Score'] = maybes_score
        
    # get possibles score if they made the cut
    possibles_index = 2
    possibles_score = df.loc[df['NAME'] == (picks[i][possibles_index])]['LATEST_SCORE'].values[0]
    position = df.loc[df['NAME'] == (picks[i][possibles_index])]['POS'].values[0]
    if position == 'CUT':
        output['Possibles Score'] = 'CUT'
    elif position == 'DQ':
        output['Possibles Score'] = 'DQ'
    else:
        output['Possibles Score'] = possibles_score
    
    # get cut score if they made the cut
    cut_index = 3
    cut_score = df.loc[df['NAME'] == (picks[i][cut_index])]['LATEST_SCORE'].values[0]
    position = df.loc[df['NAME'] == (picks[i][cut_index])]['POS'].values[0]
    if position == 'CUT':
        output['Cut Score'] = 'CUT'
    elif position == 'DQ':
        output['Cut Score'] = 'DQ'
    else:
        output['Cut Score'] = cut_score
    
    # get doubt it score if they made the cut
    doubt_it_index = 4
    doubt_it_score = df.loc[df['NAME'] == (picks[i][doubt_it_index])]['LATEST_SCORE'].values[0]
    position = df.loc[df['NAME'] == (picks[i][doubt_it_index])]['POS'].values[0]
    if position == 'CUT':
        output['Doubt It Score'] = 'CUT'
    elif position == 'DQ':
        output['Doubt It Score'] = 'DQ'
    else:
        output['Doubt It Score'] = doubt_it_score
    
    # get past it score if they made the cut
    past_it_index = 5
    past_it_score = df.loc[df['NAME'] == (picks[i][past_it_index])]['LATEST_SCORE'].values[0]
    position = df.loc[df['NAME'] == (picks[i][past_it_index])]['POS'].values[0]
    if position == 'CUT':
        output['Past It Score'] = 'CUT'
    elif position == 'DQ':
        output['Past It Score'] = 'DQ'
    else:
        output['Past It Score'] = past_it_score
        
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
                                'Past It', 'Past It Score'
                                ]].sort_values(by=['Lowest 3 Scores', 'Players making cut'], ascending=[True, False]).reset_index(drop=True)


# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Test Leaderboard.xlsx', engine='xlsxwriter')

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
for i in range(3, 16, 2):
    worksheet.set_column(i, i, 14, centre_cell_format)

# left align the player name columns
left_cell_format = workbook.add_format()
left_cell_format.set_align('left')
left_cell_format.set_font_size(14)

worksheet.set_column(1, 1, 16, left_cell_format)
for i in range(4, 15, 2):
    worksheet.set_column(i, i, 16, left_cell_format)
    

writer.save()