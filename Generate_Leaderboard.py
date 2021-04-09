import pandas as pd
import ssl
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
    'Barry Fitzpatrick': ['Jordan Spieth', 'Matthew Fitzpatrick', 'Gary Woodland', 'Jimmy Walker', 'Bernhard Langer', 'Jim Herman'],
    'Brian Bobbett': ['Justin Thomas', 'Matthew Fitzpatrick', 'Francesco Molinari', 'Ryan Palmer', 'Charl Schwartzel', 'Robert Streb'],
    'Alan Holmes': ['Justin Thomas', 'Daniel Berger', 'Billy Horschel', 'Gary Woodland', 'Martin Laird', 'Vijay Singh'],
    'Rory Collins': ['Jordan Spieth', 'Daniel Berger', 'Brian Harman', 'Robert MacIntyre', 'Charl Schwartzel', 'Vijay Singh'],
    'Alex Marshall': ['Justin Thomas', 'Webb Simpson', 'Brian Harman', 'Sebastian Munoz', 'C.T. Pan', 'Charles Osborne'],
    'Ryan Johnson': ['Dustin Johnson', 'Sergio Garcia', 'Justin Rose', 'Phil Mickelson', 'Charl Schwartzel', 'Vijay Singh'],
    'James Browne': ['Jordan Spieth', 'Matthew Fitzpatrick', 'Matt Wallace', 'Phil Mickelson', 'Jimmy Walker', 'Jose Maria Olazabal'],
    'Faolan Crowe': ['Jon Rahm', 'Webb Simpson', 'Matthew Wolff', 'Phil Mickelson', 'C.T. Pan', 'Mike Weir'],
    'Kevin Moran': ['Xander Schauffele', 'Sergio Garcia', 'Matt Kuchar', 'Christiaan Bezuidenhout', 'C.T. Pan', 'Tyler Strafaci'],
    'Brian Moran': ['Collin Morikawa', 'Tyrrell Hatton', 'Si Woo Kim', 'Gary Woodland', 'C.T. Pan', 'Tyler Strafaci'],
    'Fergal Moran': ['Jordan Spieth', 'Scottie Scheffler', 'Matt Kuchar', 'Dylan Frittelli', 'Martin Laird', 'Jose Maria Olazabal'],
    'Jack Dignam': ['Brooks Koepka', 'Shane Lowry', 'Billy Horschel', 'Zach Johnson', 'Henrik Stenson', 'Jose Maria Olazabal'],
    'Ted ODonoghue': ['Lee Westwood', 'Corey Conners', 'Matthew Wolff', 'Phil Mickelson', 'C.T. Pan', 'Robert Streb'],
    'Jack Byrne': ['Jon Rahm', 'Adam Scott', 'Billy Horschel', 'Sebastian Munoz', 'Charl Schwartzel', 'Jose Maria Olazabal'],
    'Jamie Murphy': ['Jordan Spieth', 'Sungjae Im', 'Billy Horschel', 'Kevin Kisner', 'Fred Couples', 'Mike Weir'],
    'Dave Bobbett': ['Bryson DeChambeau', 'Daniel Berger', 'Billy Horschel', 'Gary Woodland', 'Charl Schwartzel', 'Vijay Singh'],
    'Dylan Walsh': ['Justin Thomas', 'Shane Lowry', 'Marc Leishman', 'Phil Mickelson', 'Brendon Todd', 'Vijay Singh'],
    'Alan Horgan': ['Jordan Spieth', 'Daniel Berger', 'Si Woo Kim', 'Dylan Frittelli', 'Bernd Wiesberger', 'Jose Maria Olazabal'],
    'Fiachra Keane': ['Justin Thomas', 'Corey Conners', 'Francesco Molinari', 'Phil Mickelson', 'Charl Schwartzel', 'Sandy Lyle'],
    'Mark Kirwan': ['Dustin Johnson', 'Sungjae Im', 'Abraham Ancer', 'Phil Mickelson', 'Jimmy Walker', 'Tyler Strafaci'],
    'Niall Bobbett': ['Jordan Spieth', 'Daniel Berger', 'Matt Kuchar', 'Ryan Palmer', 'Bernd Wiesberger', 'Robert Streb'],
    'Shay Moran': ['Jordan Spieth', 'Sergio Garcia', 'Brian Harman', 'Phil Mickelson', 'Martin Laird', 'Vijay Singh']
}


output_list = []
# Get everyones current standings
for i in picks.keys():
    output = {
        'Player': i,
        
        'Favourites': picks[i][0],
        'Favourites Score': df.loc[df['NAME'] == (picks[i][0])]['LATEST_SCORE'].values[0],
        
        'Maybes': picks[i][1],
        'Maybes Score': df.loc[df['NAME'] == (picks[i][1])]['LATEST_SCORE'].values[0],
        
        'Possibles': picks[i][2],
        'Possibles Score': df.loc[df['NAME'] == (picks[i][2])]['LATEST_SCORE'].values[0],
        
        'Cut': picks[i][3],
        'Cut Score': df.loc[df['NAME'] == (picks[i][3])]['LATEST_SCORE'].values[0],
        
        'Doubt It': picks[i][4],
        'Doubt It Score': df.loc[df['NAME'] == (picks[i][4])]['LATEST_SCORE'].values[0],
        
        'Past It': picks[i][5],
        'Past It Score': df.loc[df['NAME'] == (picks[i][5])]['LATEST_SCORE'].values[0],
        
        'Lowest 3 Scores': df.loc[df['NAME'].isin(picks[i])].iloc[:3]['LATEST_SCORE'].sum()
    }
#     print(output)
    output_list.append(output)
    
output_df = pd.DataFrame(output_list)

# create dataframe
sorted_leaderboard = output_df[['Player',
                                'Lowest 3 Scores',
                                'Favourites','Favourites Score', 
                                'Maybes', 'Maybes Score', 
                                'Possibles', 'Possibles Score',
                                'Cut', 'Cut Score',
                                'Doubt It', 'Doubt It Score',
                                'Past It', 'Past It Score'
                                ]].sort_values(by=['Lowest 3 Scores', 'Favourites Score', 'Maybes Score', 'Possibles Score'], ascending=True).reset_index(drop=True)

print(sorted_leaderboard)