import random
import pandas as pd
import copy

import random
import copy


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


def simulate_swiss_rounds(participants, rounds=5):
    """Simulates all Swiss rounds and returns standings with detailed results."""
    standings = {participant: {'wins': 0, 'losses': 0, 'matches': []} for participant in participants}
    detailed_rounds_results = []

    for round_number in range(1, rounds + 1):
        # Sort the participants before passing to the simulate_swiss_round
        sorted_participants = sorted(standings.keys(), key=lambda x: (-standings[x]['wins'], standings[x]['losses']))
        round_results = simulate_swiss_round(sorted_participants, round_number, standings)

        # Track the participants who have already played in this round to avoid duplication
        participants_played = set()

        for match in round_results:
            if match[1] == 'Bye':  # Handle bye
                participant = match[0]
                standings[participant]['wins'] += 1
                standings[participant]['matches'].append((round_number, 'Bye', 'Win'))
            else:
                winner, loser, _ = match
                # Only add the match if neither participant has played yet this round
                if winner not in participants_played and loser not in participants_played:
                    standings[winner]['wins'] += 1
                    standings[loser]['losses'] += 1
                    standings[winner]['matches'].append((round_number, loser, 'Win'))
                    standings[loser]['matches'].append((round_number, winner, 'Loss'))
                    participants_played.update([winner, loser])

        detailed_rounds_results.append((round_number, copy.deepcopy(standings)))

    return standings, detailed_rounds_results


def export_to_excel(standings, detailed_rounds_results, de_rounds, champion):
    with pd.ExcelWriter("tournament_results.xlsx", engine='openpyxl') as writer:
        # Iterate through each Swiss round's results and rankings
        for round_number, round_standings in detailed_rounds_results:
            round_data = []
            # Track participants already added to avoid duplicating matches
            participants_added = set()

            for participant, info in round_standings.items():
                for match in info['matches']:
                    if match[0] == round_number and participant not in participants_added:  # Check the round and if we've already added this participant
                        opponent, outcome = match[1], match[2]
                        if outcome == 'Win':  # Only add the match if the participant won
                            round_data.append([participant, opponent, outcome])
                            participants_added.add(opponent)  # Add the opponent to the set so we don't add this match again
                        elif outcome == 'Bye':  # Handle bye
                            round_data.append([participant, 'Bye', 'Win'])

            df_round = pd.DataFrame(round_data, columns=['Participant', 'Opponent', 'Result'])

            # Prepare rankings data
            rankings = [[participant, info['wins'], info['losses']] for participant, info in round_standings.items()]
            rankings.sort(key=lambda x: (-x[1], x[2]))  # Sort by wins, then losses
            df_rankings = pd.DataFrame(rankings, columns=['Ranking Participant', 'Wins', 'Losses'])

            # Combine round results and rankings side by side
            combined_df = pd.concat([df_round, df_rankings], axis=1)
            combined_df.to_excel(writer, sheet_name=f'Swiss Round {round_number}', index=False)

            # Set column width to 25 for all columns
            workbook = writer.book
            worksheet = writer.sheets[f'Swiss Round {round_number}']
            for col in worksheet.columns:
                worksheet.column_dimensions[col[0].column_letter].width = 25

        # Export DE Stage Results
        for round_index, de_round in enumerate(de_rounds, start=1):
            de_data = [[match[0], match[1], match[2]] for match in de_round]
            pd.DataFrame(de_data, columns=['Participant 1', 'Participant 2', 'Winner']).to_excel(writer,
                                                                                                 sheet_name=f'DE Round {round_index}',
                                                                                                 index=False)
            # Set column width to 25
            worksheet = writer.sheets[f'DE Round {round_index}']
            for col in worksheet.columns:
                worksheet.column_dimensions[col[0].column_letter].width = 25

        # Champion
        df_champion = pd.DataFrame([{'Champion': champion}])
        df_champion.to_excel(writer, sheet_name='Champion', index=False)
        # Set column width to 25 for the Champion sheet
        worksheet = writer.sheets['Champion']
        for col in worksheet.columns:
            worksheet.column_dimensions[col[0].column_letter].width = 25


# Ensure the rest of your functions are defined here, particularly those for the DE stage.

def qualify_for_de(sorted_standings, top_n=16):
    """Selects the top_n participants based on Swiss stage performance for the DE stage."""
    # No need to sort here since it's already sorted before passing
    return sorted_standings[:top_n]


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


def main():
    # Predefined list of participants
    participants = ["Toni", "Stoyan", "Plamen", "Bobi", "Petyo", "Rosko", "Sasho", "Marto", "Nelly", "Nati", "Alexi",
                    "Tsveti", "Misho", "Pesho", "Alex", "Sasho M", "Reni", "Miro", "Gabi", "Geri", "Didi", "Kalata",
                    "Yavkata", "Ivo", "Marto S"]

    print("\nSimulating Swiss Stage...")
    standings, detailed_rounds_results = simulate_swiss_rounds(participants, rounds=4)

    # Sort standings to determine top participants for DE
    sorted_standings = sorted(standings.items(), key=lambda x: (x[1]['wins'],
                                                                -x[1]['losses'],
                                                                sum(1 for match in x[1]['matches'] if
                                                                    match[2] == 'Win'),
                                                                -sum(1 for match in x[1]['matches'] if
                                                                     match[2] == 'Loss')),
                              reverse=True)

    # Qualifiers for DE Stage
    qualifiers = qualify_for_de(sorted_standings, top_n=16)
    print("\nQualifiers for DE Stage:")
    for idx, (participant, _) in enumerate(qualifiers, 1):
        print(f"{idx}. {participant}")

    # Simulating DE Stage
    de_participants = [qualifier[0] for qualifier in qualifiers]
    champion, de_rounds = simulate_de(de_participants)

    print(f"\nChampion: {champion}")

    # Export results to Excel
    export_to_excel(standings, detailed_rounds_results, de_rounds, champion)
    print("Tournament results have been exported to 'tournament_results.xlsx'.")


if __name__ == "__main__":
    main()
