import pandas as pd
import random
import copy
import os
from openpyxl import load_workbook
import pdb;

def generate_initial_pairings(participants):
    """Generates pairings for a round in a Swiss-system tournament."""
    # Shuffle only for the first round, subsequent rounds should be based on rankings
    # random.shuffle(participants)  # No longer needed as sorting is done in simulate_swiss_round
    # Pair participants, ensuring each is only paired once
    return [participants[i:i + 2] for i in range(0, len(participants), 2)]


def simulate_swiss_round(participants, round_number, previous_rounds_results):
    """Simulates a single Swiss round and updates standings"""
    pairings = generate_initial_pairings(participants)

    round_results = []
    for pair in pairings:
        if len(pair) == 1:  # Handle bye if an odd number of participants
            round_results.append((pair[0], 'Bye', 'Win'))
        else:
            winner = random.choice(pair)
            loser = pair[1] if pair[0] == winner else pair[0]
            round_results.append((winner, loser, 'Win'))  # Only record winner's perspective
    return round_results


def simulate_swiss_rounds(participants, num_rounds=1, start_round=1, initial_standings=None):
    if initial_standings is None:
        standings = {participant: {'wins': 0, 'losses': 0, 'matches': []} for participant in participants}
    else:
        standings = initial_standings

    detailed_rounds_results = []

    for participant in participants:
        if participant not in standings:
            standings[participant] = {'wins': 0, 'losses': 0, 'matches': []}

    for round_number in range(start_round, start_round + num_rounds):
        sorted_participants = sorted(standings.keys(), key=lambda x: (-standings[x]['wins'], standings[x]['losses'], x))
        round_results = simulate_swiss_round(sorted_participants, round_number, standings)

        for result in round_results:
            if result[1] == 'Bye':
                participant = result[0]
                standings[participant]['wins'] += 1
                standings[participant]['matches'].append((round_number, 'Bye', 'Win'))
            else:
                winner, loser, _ = result
                standings[winner]['wins'] += 1
                standings[loser]['losses'] += 1
                standings[winner]['matches'].append((round_number, loser, 'Win'))
                standings[loser]['matches'].append((round_number, winner, 'Loss'))

        detailed_rounds_results.append((round_number, copy.deepcopy(standings)))

    return standings, detailed_rounds_results

def generate_pairings_based_on_rankings(participants):
    """
    Generates pairings for the next round based on current rankings.
    Assumes 'participants' is a list sorted by rankings, highest first.
    """
    pairings = []
    # Generate pairings by taking two at a time from the sorted list
    for i in range(0, len(participants), 2):
        # Check if there's an odd participant out for a bye
        if i + 1 < len(participants):
            pair = (participants[i], participants[i+1], "")
            pairings.append(pair)
        else:
            # Assign a bye (win) if odd number of participants
            pair = (participants[i], 'Bye', 'Win')
            pairings.append(pair)
    return pairings


def de_generate_pairings_based_on_rankings(participants):
    pairings = []
    num_participants = len(participants)
    for i in range(num_participants // 2):
        pairings.append((participants[i], participants[num_participants - i - 1], ""))
    return pairings
    

def export_next_round_to_excel(filename, standings, round_number):
    # Load the existing workbook
    book = load_workbook(filename)

    # Ensure there's at least one visible sheet
    if all(ws.sheet_state == 'hidden' for ws in book.worksheets):
        book.worksheets[0].sheet_state = 'visible'

    # Calculate rankings based on standings
    rankings_data = []
    for participant, info in standings.items():
        rankings_data.append([participant, info['wins'], info['losses']])

    # Sort rankings data based on wins, then losses
    rankings_data.sort(key=lambda x: (-x[1], x[2]))  # Sort by wins, then losses
    df_rankings = pd.DataFrame(rankings_data, columns=['Standings', 'Wins', 'Losses']) #this should go to the first sheet only
    
    # Define the threshold for the start of the DE stage
    DE_THRESHOLD = 4
    
    # From here we need to check if the round_number is bigger than 4 and if so we need to start the DE stage
    if round_number > DE_THRESHOLD: #this tells us that we are into the DE stage
        sorted_standings = sorted(standings.items(), key=lambda x: (-x[1]['wins'], x[1]['losses']))


        #TODO this part for DE is far from ready
        # Get the participants from the previous DE stage and if this is the first DE stage than take the top 8 participants
        if round_number == DE_THRESHOLD+1:
            top_n_participants = qualify_for_de(sorted_standings, top_n=8)
        else:
            top_n_participants = de_read_last_round_and_update_standings(filename)
            
        # Export DE results to Excel
        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            if(round_number <= DE_THRESHOLD+1):
                existing_data = pd.read_excel(filename, sheet_name=f'Swiss Round {round_number-1}', usecols=[0, 1, 2])
                combined_data = pd.concat([existing_data, df_rankings], axis=1)
                combined_data.to_excel(writer, sheet_name=f'Swiss Round {round_number-1}', index=False, startcol=0)
            
            # Generate pairings for the next round based on current rankings
            de_pairings = de_generate_pairings_based_on_rankings(top_n_participants)

            # Convert pairings to DataFrame for easier manipulation and export
            df_de_round = pd.DataFrame(de_pairings, columns=['Participant', 'Opponent',"Result"])

            # Export DE round results to Excel
            df_de_round.to_excel(writer, sheet_name=f'DE Round {round_number}', index=False)
            
            #This code just sets the width of the columns to 25
            worksheet = writer.sheets[f'DE Round {round_number}']
            for col in worksheet.columns:
                worksheet.column_dimensions[col[0].column_letter].width = 25
                
            if(round_number <= DE_THRESHOLD+1):
                worksheet = writer.sheets[f'Swiss Round {round_number-1}']
                for col in worksheet.columns:
                    worksheet.column_dimensions[col[0].column_letter].width = 25
            
    else:
        # Since you want to prepare for the next round without results, we'll use the rankings to generate pairings
        # Assuming you have a function to generate pairings from rankings or standings
        participants_sorted_by_rankings = [participant for participant, _ in
                                           sorted(standings.items(), key=lambda x: (-x[1]['wins'], x[1]['losses']))]
        round_data = generate_pairings_based_on_rankings(participants_sorted_by_rankings)
        # Convert round_data to DataFrame for easier manipulation and export
        df_next_round = pd.DataFrame(round_data, columns=['Participant', 'Opponent', 'Result'])
        next_round_sheet_name = f'Swiss Round {round_number}'
            
        print(f"{next_round_sheet_name} has been added to 'tournament_results.xlsx'.")
        # Use pandas to write DataFrame to the Excel sheet in the workbook
        with pd.ExcelWriter(filename, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            df_next_round.to_excel(writer, sheet_name=next_round_sheet_name, index=False)
            
            # Get the existing data in columns 1 and 2
            existing_data = pd.read_excel(filename, sheet_name=f'Swiss Round {round_number-1}', usecols=[0, 1, 2])
            combined_data = pd.concat([existing_data, df_rankings], axis=1)
            combined_data.to_excel(writer, sheet_name=f'Swiss Round {round_number-1}', index=False, startcol=0)
            
            #This code just sets the width of the columns to 25
            worksheet = writer.sheets[next_round_sheet_name]
            for col in worksheet.columns:
                worksheet.column_dimensions[col[0].column_letter].width = 25
                
            worksheet = writer.sheets[f'Swiss Round {round_number-1}']
            for col in worksheet.columns:
                worksheet.column_dimensions[col[0].column_letter].width = 25


def read_last_round_and_update_standings(filename, initial_standings):
    xls = pd.ExcelFile(filename)
    if(xls.sheet_names.__len__ == 1):
        latest_round_sheet_name = xls.sheet_names[xls.sheet_names[0]]
        previous_round_sheet_name = xls.sheet_names[xls.sheet_names[0]]
    else:
        latest_round_sheet_name = xls.sheet_names[xls.sheet_names.__len__() - 1]
        previous_round_sheet_name = xls.sheet_names[xls.sheet_names.__len__() - 2]
    df_latest_round = pd.read_excel(filename, sheet_name=latest_round_sheet_name)
    df_previous_round = pd.read_excel(filename, sheet_name=previous_round_sheet_name)

    try:
        last_round_played = int(latest_round_sheet_name.split()[-1])
    except ValueError:
        print(f"Error extracting round number from sheet name: {latest_round_sheet_name}")
        last_round_played = 0

  # Check if standings exist in the previous file
    if 'Wins' in df_previous_round.columns:
        standings = {row['Standings']: {'wins': row['Wins'], 'losses': row['Losses'], 'matches': []} for index, row in df_previous_round.iterrows()}
    else:
        standings = initial_standings

    for index, row in df_latest_round.iterrows():
        participant = row['Participant']
        opponent = row['Opponent']
        result = row['Result']  # Assumes 'Win', 'Loss', or 'Bye'
           
        if opponent.lower() == 'bye':
            standings[participant]['wins'] += 1
            standings[participant]['matches'].append((last_round_played, opponent, result))
        elif result.lower() == 'win':
            standings[participant]['wins'] += 1
            standings[opponent]['losses'] += 1
            standings[participant]['matches'].append((last_round_played, opponent, result))
            standings[opponent]['matches'].append((last_round_played, participant, 'Loss'))
        elif result.lower() == 'loss':
            standings[participant]['losses'] += 1
            standings[opponent]['wins'] += 1
            standings[participant]['matches'].append((last_round_played, opponent, result))
            standings[opponent]['matches'].append((last_round_played, participant, 'Win'))

    return standings, last_round_played


def de_read_last_round_and_update_standings(filename):
    xls = pd.ExcelFile(filename)
    if(xls.sheet_names.__len__ == 1):
        latest_round_sheet_name = xls.sheet_names[xls.sheet_names[0]]
    else:
        latest_round_sheet_name = xls.sheet_names[xls.sheet_names.__len__() - 1]
    df_latest_round = pd.read_excel(filename, sheet_name=latest_round_sheet_name)

    winners = []
    for index, row in df_latest_round.iterrows():
        participant = row['Participant']
        opponent = row['Opponent']
        result = row['Result']  # Assumes 'Win', 'Loss', or 'Bye'
           
        if result.lower() == 'win':
            winners.append(participant)
        elif result.lower() == 'loss':
            winners.append(opponent)

    return winners


# Ensure the rest of your functions are defined here, particularly those for the DE stage.


def qualify_for_de(sorted_standings, top_n=8):
    """Selects the top_n participants based on Swiss stage performance for the DE stage."""
    # No need to sort here since it's already sorted before passing
    
    sorted_standings_names = [participant for participant, _ in sorted_standings]
    return sorted_standings_names[:top_n]
    # return sorted_standings[:top_n]


def simulate_de(participants):

    """Simulates the Direct Elimination (DE) stage and tracks the matchups and winners."""
    de_rounds = []
    current_participants = participants  # Assuming 'participants' is already a list of participant names

    while len(current_participants) > 1:
        next_round_participants = []
        round_matches = []
        for i in range(0, len(current_participants), 2):
            if i + 1 < len(current_participants):  # Ensure there is a pair to compete
                p1 = current_participants[i]
                p2 = current_participants[i + 1]
                winner = random.choice([p1, p2])
                next_round_participants.append(winner)
                round_matches.append((p1, p2, winner))
        if round_matches:  # Only add to de_rounds if there were matches
            de_rounds.append(round_matches)
        current_participants = next_round_participants

    champion = current_participants[0] if current_participants else None
    return champion, de_rounds


def create_initial_pairings(participants):
    # Shuffle the participants
    random.shuffle(participants)

    # Pair them up
    pairings = list(zip(participants[::2], participants[1::2]))

    # If there's an odd number of participants, add a 'Bye' pairing for the last one
    if len(participants) % 2 != 0:
        pairings.append((participants[-1], 'Bye'))

    return pairings

def export_to_excel(pairings):
    # Prepare data for export
        round_data = [[pair[0], pair[1], ''] for pair in pairings]
        df_round = pd.DataFrame(round_data, columns=['Participant', 'Opponent', 'Result'])

        # Export to Excel
        with pd.ExcelWriter("tournament_results.xlsx", engine='openpyxl') as writer:
            df_round.to_excel(writer, sheet_name='Swiss Round 1', index=False)

            worksheet = writer.sheets[f'Swiss Round 1']
            for col in worksheet.columns:
                worksheet.column_dimensions[col[0].column_letter].width = 25
        print("Initial round has been exported to 'tournament_results.xlsx'.")

def main():
    excel_filename = "tournament_results.xlsx"
    participants = ["Чом", "Жеко", "Рени", "Алекс", "Марто С.", "Миро", "Цвети", "Диди", "Нати З.",
                    "Роско", "Сандо", "Явката", "Стоян", "Нели", "Пламен", "Петьо", "Алекси", "Стан", 
                    "Калата К.", "Нати Т.", "Александър К.", "Теодор Й.", "Габи"]

    # Initialize standings
    standings = {participant: {'wins': 0, 'losses': 0, 'matches': []} for participant in participants}

    if os.path.exists(excel_filename):
        print("\nFile exists. Reading last round results and updating standings...")
        # This function needs to be implemented to read the last round results and update standings.
        # Assuming it returns the updated standings and the round number of the last round played.
        standings, last_round_played = read_last_round_and_update_standings(excel_filename, standings)

        # Export next round's pairings to Excel
        # This function needs to append the new round's pairings to the existing Excel file.
        export_next_round_to_excel(excel_filename, standings, last_round_played + 1)
        print(f"Round {last_round_played}'s pairings have been added to 'tournament_results.xlsx'.")
    else:
        print("\nGenerating Initial Swiss Stage...")
        # Create initial pairings
        pairings = create_initial_pairings(participants)
        export_to_excel(pairings)

if __name__ == "__main__":
    main()
