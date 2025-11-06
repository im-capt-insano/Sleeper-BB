import pandas as pd
import gspread
import time
import math
from sleeper_wrapper import League, Players
from datetime import datetime as dt
import nfl_data_py as nfl
import openpyxl as pyxl
from gspread_formatting import *
from gspread.exceptions import APIError

class Dumpster_Dynasty:

    def __init__(self, *args):
        self.args = args        
        self.Sheet_url = 'https://docs.google.com/spreadsheets/d/1ko1XnttApOFkA1x1RF95Ls5C-aM98p766ht058Il350/edit?pli=1&gid=962080300#gid=962080300'
        self.Service_account = r'C:\Users\Jed\AppData\Local\Programs\Python\Python310\Lib\site-packages\gspread\dumpster-dynasty-bb.json'
        self.Credentials = {"installed":{"client_id":"215364173021-j261liv1nscj8tj4es9do2mun9t4a603.apps.googleusercontent.com","project_id":"dumpster-dynasty-bb","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
        self.Start_year = 2024
        starting_pos = []
        self.Num_qb = 1
        for r in range(self.Num_qb):
            starting_pos.append('QB')
        self.Num_rb = 2
        for r in range(self.Num_rb):
            starting_pos.append('RB')
        self.Num_wr = 2
        for r in range(self.Num_wr):
            starting_pos.append('WR')
        self.Num_te = 1
        for r in range(self.Num_te):
            starting_pos.append('TE')
        self.Num_flex = 2
        for r in range(self.Num_flex):
            starting_pos.append('FLEX')
        self.Num_sflex = 1
        for r in range(self.Num_sflex):
            starting_pos.append('SFLEX')
        self.Num_k = 1
        for r in range(self.Num_k):
            starting_pos.append('K')
        self.Num_dst = 1
        for r in range(self.Num_dst):
            starting_pos.append('DST')
        self.Starting_pos = starting_pos
        self.Stat_pts = self.Stat_points()
        self.Num_starter = len(starting_pos)
        self.Num_bench = 14
        self.Num_ir = 3
        self.Num_taxi = 3
        self.Roster_size = self.Num_starter + self.Num_bench + self.Num_ir + self.Num_taxi
        self.Num_playoff_teams = 6
        self.Num_bye_teams = 2

    def Update(self, input, *args):
        if input >= self.Start_year:
            match input:
                case 2024:
                    self.League = League(1128506799759429632)
                case 2025:
                    self.League = League(1180265469580161024)
            self.playoff_quarter = [15]
            self.playoff_semi = [16]
            self.playoff_final = [17]
            self.Name = self.League.get_league_name()
            self.Rosters = self.League.get_rosters()
            all_users = self.League.get_users()
            user_info = []
            for owner in range(len(all_users)):
                user_info.append(list(map(all_users[owner].get, ['user_id', 'display_name'])))
            user_name = []
            for roster in range(len(self.Rosters)):
                need = self.Rosters[roster]['owner_id']
                for user in user_info:
                    if user[0] == need:
                        user_name.append(user)
                        break
            self.Users = user_name
            self.Standings = self.League.get_standings(self.Rosters, all_users)
        else:
            self.Matchups = self.League.get_matchups(input)

    def Stat_points(self):
        Stat_pts = []
        Stat_pts.append(['pass_yd', 0.04])
        Stat_pts.append(['pass_td', 4])
        Stat_pts.append(['pass_conv2', 2])
        Stat_pts.append(['pass_int', -1])
        Stat_pts.append(['rush_yd', 0.1])
        Stat_pts.append(['rush_td', 6])
        Stat_pts.append(['rush_conv2', 2])
        Stat_pts.append(['rec', 0.5])
        Stat_pts.append(['rec_yd', 0.1])
        Stat_pts.append(['rec_td', 6])
        Stat_pts.append(['rec_conv2', 2])
        Stat_pts.append(['fg_yd', 0.1])
        Stat_pts.append(['pat_make', 1])
        Stat_pts.append(['fg_miss_0_19', -4])
        Stat_pts.append(['fg_miss_20_29', -3])
        Stat_pts.append(['fg_miss_30_39', -2])
        Stat_pts.append(['fg_miss_40_49', -1])
        Stat_pts.append(['fg_miss_50_59', -1])
        Stat_pts.append(['pat_miss', -1])
        Stat_pts.append(['def_td', 6])
        Stat_pts.append(['def_allow_0', 10])
        Stat_pts.append(['def_allow_0_6', 7])
        Stat_pts.append(['def_allow_7_13', 4])
        Stat_pts.append(['def_allow_14_20', 1])
        Stat_pts.append(['def_allow_21_27', 0])
        Stat_pts.append(['def_allow_28_34', -1])
        Stat_pts.append(['def_allow_35', -4])
        Stat_pts.append(['4th_stop', 0])
        Stat_pts.append(['sack', 1])
        Stat_pts.append(['pick', 2])
        Stat_pts.append(['f_recover', 2])
        Stat_pts.append(['safe', 2])
        Stat_pts.append(['fumble_force', 1])
        Stat_pts.append(['block_kick', 2])
        Stat_pts.append(['st_td', 6])
        Stat_pts.append(['st_ff', 1])
        Stat_pts.append(['st_fr', 1])
        Stat_pts.append(['fumble_loss', -2])
        Stat_pts.append(['fumble_td', 6])
        Stat_pts = pd.DataFrame(Stat_pts, columns=['Stat', 'Points'])
        return Stat_pts

class RIP_Harambe:

    def __init__(self, *args):
        self.args = args
        self.Sheet_url = 'https://docs.google.com/spreadsheets/d/1uy91bw37lcMzz2gcw7UDDgI6NJmkR887EEvmDsJHfjc/edit?gid=0#gid=0'
        self.Service_account = r'C:\Users\Jed\AppData\Local\Programs\Python\Python310\Lib\site-packages\gspread\dumpster-dynasty-bb.json'
        self.Credentials = {"installed":{"client_id":"215364173021-j261liv1nscj8tj4es9do2mun9t4a603.apps.googleusercontent.com","project_id":"dumpster-dynasty-bb","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
        self.Start_year = 2023
        starting_pos = []
        self.Num_qb = 1
        for r in range(self.Num_qb):
            starting_pos.append('QB')
        self.Num_rb = 2
        for r in range(self.Num_rb):
            starting_pos.append('RB')
        self.Num_wr = 2
        for r in range(self.Num_wr):
            starting_pos.append('WR')
        self.Num_te = 1
        for r in range(self.Num_te):
            starting_pos.append('TE')
        self.Num_flex = 2
        for r in range(self.Num_flex):
            starting_pos.append('FLEX')
        self.Num_sflex = 0
        for r in range(self.Num_sflex):
            starting_pos.append('SFLEX')
        self.Num_k = 1
        for r in range(self.Num_k):
            starting_pos.append('K')
        self.Num_dst = 1
        for r in range(self.Num_dst):
            starting_pos.append('DST')
        self.Starting_pos = starting_pos
        self.Num_starter = len(starting_pos)
        self.Num_bench = 16
        self.Num_ir = 2
        self.Num_taxi = 3
        self.Roster_size = self.Num_starter + self.Num_bench + self.Num_ir + self.Num_taxi
        self.Num_playoff_teams = 7
        self.Num_bye_teams = 1

    def Update(self, input, *args):
        if input >= self.Start_year:
            match input:
                case 2023:
                    self.League = League(919311157205319680)
                case 2024:
                    self.League = League(1088873848449097728)
                case 2025:
                    self.League = League(1180228760887463936)
            self.playoff_quarter = [14]
            self.playoff_semi = [15]
            self.playoff_final = [16, 17]
            self.Name = self.League.get_league_name()
            self.Rosters = self.League.get_rosters()
            all_users = self.League.get_users()
            user_info = []
            for owner in range(len(all_users)):
                user_info.append(list(map(all_users[owner].get, ['user_id', 'display_name'])))
            user_name = []
            for roster in range(len(self.Rosters)):
                need = self.Rosters[roster]['owner_id']
                for user in user_info:
                    if user[0] == need:
                        user_name.append(user)
                        break
            self.Users = user_name
            self.Standings = self.League.get_standings(self.Rosters, all_users)
        else:
            self.Matchups = self.League.get_matchups(input)

    def Stat_points(self):
            Stat_pts = []
            Stat_pts.append(['pass_yd', 0.04])
            Stat_pts.append(['pass_td', 4])
            Stat_pts.append(['pass_conv2', 2])
            Stat_pts.append(['pass_int', -1])
            Stat_pts.append(['rush_yd', 0.1])
            Stat_pts.append(['rush_td', 6])
            Stat_pts.append(['rush_conv2', 2])
            Stat_pts.append(['rec', 1])
            Stat_pts.append(['rec_yd', 0.1])
            Stat_pts.append(['rec_td', 6])
            Stat_pts.append(['rec_conv2', 2])
            Stat_pts.append(['fg_yd', 0.1])
            Stat_pts.append(['pat_make', 1])
            Stat_pts.append(['fg_miss_0_19', 0])
            Stat_pts.append(['fg_miss_20_29', 0])
            Stat_pts.append(['fg_miss_30_39', 0])
            Stat_pts.append(['fg_miss_40_49', 0])
            Stat_pts.append(['fg_miss_50_59', 0])
            Stat_pts.append(['pat_miss', -1])
            Stat_pts.append(['def_td', 6])
            Stat_pts.append(['def_allow_0', 10])
            Stat_pts.append(['def_allow_0_6', 7])
            Stat_pts.append(['def_allow_7_13', 4])
            Stat_pts.append(['def_allow_14_20', 1])
            Stat_pts.append(['def_allow_21_27', 0])
            Stat_pts.append(['def_allow_28_34', -1])
            Stat_pts.append(['def_allow_35', -4])
            Stat_pts.append(['4th_stop', 1])
            Stat_pts.append(['sack', 1])
            Stat_pts.append(['pick', 2])
            Stat_pts.append(['f_recover', 2])
            Stat_pts.append(['safe', 2])
            Stat_pts.append(['fumble_force', 1])
            Stat_pts.append(['block_kick', 2])
            Stat_pts.append(['st_td', 6])
            Stat_pts.append(['st_ff', 1])
            Stat_pts.append(['st_fr', 1])
            Stat_pts.append(['fumble_loss', -2])
            Stat_pts.append(['fumble_td', 6])
            Stat_pts = pd.DataFrame(Stat_pts, columns=['Stat', 'Points'])
            return Stat_pts

def int_to_column(n):
    result = ''
    while n > 0:
        n, remainder = divmod(n-1, 26)
        result = chr(65+remainder) + result
    return result

def Best_player(possible, starters, *args):
    if starters:
        starters_df = pd.DataFrame(starters, columns=possible.columns)
        possible = possible[~possible['gsis_id'].isin(starters_df['gsis_id'])]
    if not possible.empty:
        starters.append(possible.loc[possible.idxmax().Pts].values)

    return starters

def Border_format(sides):
    match sides:
        case 'none':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID'},
                'bottom': {'style': 'SOLID'},
                'left': {'style': 'SOLID'},
                'right': {'style': 'SOLID'}})
        case 't':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID_THICK'},
                'bottom': {'style': 'SOLID'},
                'left': {'style': 'SOLID'},
                'right': {'style': 'SOLID'}})
        case 'b':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID'},
                'bottom': {'style': 'SOLID_THICK'},
                'left': {'style': 'SOLID'},
                'right': {'style': 'SOLID'}})
        case 'l':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID'},
                'bottom': {'style': 'SOLID'},
                'left': {'style': 'SOLID_THICK'},
                'right': {'style': 'SOLID'}})
        case 'r':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID'},
                'bottom': {'style': 'SOLID'},
                'left': {'style': 'SOLID'},
                'right': {'style': 'SOLID_THICK'}})
        case 'tb':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID_THICK'},
                'bottom': {'style': 'SOLID_THICK'},
                'left': {'style': 'SOLID'},
                'right': {'style': 'SOLID'}})
        case 'lr':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID'},
                'bottom': {'style': 'SOLID'},
                'left': {'style': 'SOLID_THICK'},
                'right': {'style': 'SOLID_THICK'}})
        case 'tl':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID_THICK'},
                'bottom': {'style': 'SOLID'},
                'left': {'style': 'SOLID_THICK'},
                'right': {'style': 'SOLID'}})
        case 'tr':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID_THICK'},
                'bottom': {'style': 'SOLID'},
                'left': {'style': 'SOLID'},
                'right': {'style': 'SOLID_THICK'}})
        case 'bl':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID'},
                'bottom': {'style': 'SOLID_THICK'},
                'left': {'style': 'SOLID_THICK'},
                'right': {'style': 'SOLID'}})
        case 'br':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID'},
                'bottom': {'style': 'SOLID_THICK'},
                'left': {'style': 'SOLID'},
                'right': {'style': 'SOLID_THICK'}})
        case 'tlr':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID_THICK'},
                'bottom': {'style': 'SOLID'},
                'left': {'style': 'SOLID_THICK'},
                'right': {'style': 'SOLID_THICK'}})
        case 'blr':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID'},
                'bottom': {'style': 'SOLID_THICK'},
                'left': {'style': 'SOLID_THICK'},
                'right': {'style': 'SOLID_THICK'}})
        case 'tbl':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID_THICK'},
                'bottom': {'style': 'SOLID_THICK'},
                'left': {'style': 'SOLID_THICK'},
                'right': {'style': 'SOLID'}})
        case 'tbr':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID_THICK'},
                'bottom': {'style': 'SOLID_THICK'},
                'left': {'style': 'SOLID'},
                'right': {'style': 'SOLID_THICK'}})
        case 'all':
            brdr_fmt = CellFormat(borders={
                'top': {'style': 'SOLID_THICK'},
                'bottom': {'style': 'SOLID_THICK'},
                'left': {'style': 'SOLID_THICK'},
                'right': {'style': 'SOLID_THICK'}})

    return brdr_fmt

if __name__ == "__main__":
    sl_time = 3
    start_week = 1
    end_week = 18
    year_td = dt.now().year
    max_col = (7*end_week) + (end_week-1)
    DST_list = ['ARI', 'ATL', 'BAL', 'BUF', 'CAR', 'CHI', 'CIN', 'CLE', 'DAL', 'DEN', 'DET', 'GB', 'HOU', 'IND', 'JAX', 'KC', 'LAC', 'LAR', 'LV', 'MIA', 'MIN', 'NE', 'NYG', 'NYJ', 'NO', 'PHI', 'PIT', 'SEA', 'SF', 'TB', 'TEN', 'WAS']
    # gc, authorized_user = gspread.oauth_from_dict(credentials)
    Dumpster = Dumpster_Dynasty()
    Harambe = RIP_Harambe()
    leagues_all = [Dumpster, Harambe]
    git = 'https://github.com/im-capt-insano/Sleeper-BB'
    # all_players = Players().get_all_players('nfl')
    # week_data = nfl.import_weekly_data([year])
    txt_fmt_bld = CellFormat(textFormat=TextFormat(bold=True))
    txt_fmt_bld_ctr = CellFormat(textFormat=TextFormat(bold=True), horizontalAlignment='CENTER')
    red = Color(230/255, 124/255, 115/255)
    yellow = Color(255/255, 214/255, 102/255)
    green = Color(87/255, 187/255, 138/255)
    white = Color(1, 1, 1)
    for cur_league in leagues_all:
        gc = gspread.service_account(filename=cur_league.Service_account)
        gsheet = gc.open_by_url(cur_league.Sheet_url)
        # for year in range(cur_league.Start_year, year_td+1):
        for year in range(year_td, year_td+1):
            match year:
                case 2023:
                    player_table = pd.read_csv('https://github.com/nflverse/nflverse-data/releases/download/rosters/roster_2023.csv')
                case 2024:
                    player_table = pd.read_csv('https://github.com/nflverse/nflverse-data/releases/download/rosters/roster_2024.csv')
                case 2025:
                    player_table = pd.read_csv('https://github.com/nflverse/nflverse-data/releases/download/rosters/roster_2025.csv')
            cur_league.Update(year)
            num_owners = len(cur_league.Users)
            num_playoff = cur_league.Num_playoff_teams
            non_playoff = num_owners-num_playoff
            #   Year Summary Update needs to happen after player update due to formulas not working if the reference sheet does not yet exist
            # for week in range(start_week, end_week+1):
            for week in range(start_week, end_week):
                #   Grab data from each specific week
                cur_league.Update(week, year)
                #   Create variables which will be used in the for loop
                header_size = 2
                summary_size = 1
                data_num_cols = 7
                team_size = cur_league.Roster_size
                r1 = 1 + (year-cur_league.Start_year)*(1+header_size+summary_size+cur_league.Roster_size)
                r2 = r1+1
                r3 = r2+1
                r4 = r2+cur_league.Num_starter
                r5 = r4+1
                r6 = r5+1
                r7 = r5+cur_league.Roster_size-cur_league.Num_starter
                for owner in range(0, num_owners):
                    illegal = False
                    point_count = True
                    owner_matchups = cur_league.Matchups[owner]
                    if week in cur_league.playoff_quarter + cur_league.playoff_semi + cur_league.playoff_final:
                        bye = False
                        playoffs = cur_league.League.get_playoff_winners_bracket()
                        if week in cur_league.playoff_semi + cur_league.playoff_final:
                            playoff_matchups = [1,2]
                            if owner_matchups['matchup_id'] not in playoff_matchups:
                                point_count = False
                        else:
                            temp = cur_league.Num_playoff_teams-cur_league.Num_bye_teams
                            playoff_matchups = range(1, int(temp/2) + cur_league.Num_bye_teams + 1)
                            for match in range(0, len(playoffs)):
                                if playoffs[match]['t1'] == owner+1:
                                    bye = True
                                    continue
                                if playoffs[match]['t2'] == owner+1:
                                    bye = True
                                    continue
                            if owner_matchups['matchup_id'] not in playoff_matchups and not bye:
                                point_count = False
                    starters = owner_matchups['starters']
                    if starters is None:
                        continue 
                    roster = owner_matchups['players']
                    bench = list(set(roster) ^ set(starters))
                    user_id = cur_league.Users[owner][0]
                    user_name = cur_league.Users[owner][1]
                    roster_names = []
                    roster_pos = []
                    roster_data = []
                    bench_pts = []
                    starter_pts = []
                    roster_data_dict = owner_matchups['players_points']
                    #   Seperate actual starters from actual bench
                    for sleeper_id, pts in roster_data_dict.items():
                        if sleeper_id in starters:
                            starter_pts.append([sleeper_id, pts])
                        else:
                            bench_pts.append([sleeper_id, pts])
                        roster_data.append([sleeper_id, pts])
                    roster_gsis_ids = []
                    roster_sleeper_ids = []
                    #   Turn player ID's in to player names
                    for player in range(len(roster)):
                        #   Seperate out defenses
                        if str(roster[player]) in DST_list:
                            player_name = roster[player]
                            player_gsis_id = roster[player]
                            player_pos = 'DST'
                        elif roster[player] == '11539':
                            player_name = 'Jake Bates'
                            player_gsis_id = '00-0039172'
                            player_pos = 'K'
                        else:
                            #   Jake Bates
                            if roster[player] == '11539':
                                player_name = 'Jake Bates'

                            player_nfl = player_table[player_table['sleeper_id'].isin([int(roster[player])])][['full_name', 'gsis_id', 'position', 'depth_chart_position']]
                            #   Fill in data for not found players
                            try:
                                player_pos = player_nfl['position'].values[0]
                                player_name = player_nfl['full_name'].values[0]
                                player_gsis_id = player_nfl['gsis_id'].values[0]
                            except:
                                player_pos = 'NOT FOUND'
                                player_name = 'NOT FOUND'
                                player_gsis_id = 'NOT FOUND'
                            #   Set the eligible poistions for Tayson Hill
                            if player_name == 'Taysom Hill':
                                player_pos = 'TE & QB'
                        roster_data[player].append(player_name)
                        roster_data[player].append(player_pos)
                        roster_data[player].append(player_gsis_id)
                    #   Full roster
                    roster_full = pd.DataFrame(roster_data, columns=['sleeper_id', 'Pts', 'Name', 'POS', 'gsis_id'])
                    #   Actual starters
                    starters_actual = roster_full[roster_full['sleeper_id'].isin(starters)]
                    #   Actual bench
                    bench_actual = roster_full[~roster_full['sleeper_id'].isin(starters)]
                    #   Full roster
                    roster_full = roster_full.sort_values(by=['POS', 'Pts'], ascending=[True, False])
                    #   Sort actual player lists by position then points for readability
                    starters_actual = starters_actual.sort_values(by=['POS', 'Pts'], ascending=[True, False])
                    bench_actual = bench_actual.sort_values(by=['POS', 'Pts'], ascending=[True, False])
                    bench_actual = bench_actual.reset_index(drop=True)
                    #   Break out full roster to individual positions
                    roster_qb = roster_full[roster_full['POS'].str.contains('QB')]
                    roster_rb = roster_full[roster_full['POS'].str.contains('RB')]
                    roster_wr = roster_full[roster_full['POS'].str.contains('WR')]
                    roster_te = roster_full[roster_full['POS'].str.contains('TE')]
                    roster_k = roster_full[roster_full['POS'].str.contains('K')]
                    roster_dst = roster_full[roster_full['POS'].str.contains('DST')]
                    roster_flex = roster_full[roster_full['POS'].str.contains('WR')|roster_full['POS'].str.contains('RB')|roster_full['POS'].str.contains('TE')]
                    roster_sflex = roster_full[roster_full['POS'].str.contains('QB')|roster_full['POS'].str.contains('WR')|roster_full['POS'].str.contains('RB')|roster_full['POS'].str.contains('TE')]
                    starters_bb = []
                    #   Best player analysis for each position
                    for x in range(cur_league.Num_qb):
                        starters_bb = Best_player(roster_qb, starters_bb)
                    for x in range(cur_league.Num_rb):
                        starters_bb = Best_player(roster_rb, starters_bb)
                    for x in range(cur_league.Num_wr):
                        starters_bb = Best_player(roster_wr, starters_bb)
                    for x in range(cur_league.Num_te):
                        starters_bb = Best_player(roster_te, starters_bb)
                    for x in range(cur_league.Num_flex):
                        starters_bb = Best_player(roster_flex, starters_bb)
                    for x in range(cur_league.Num_sflex):
                        starters_bb = Best_player(roster_sflex, starters_bb)
                    for x in range(cur_league.Num_k):
                        starters_bb = Best_player(roster_k, starters_bb)
                    for x in range(cur_league.Num_dst):
                        starters_bb = Best_player(roster_dst, starters_bb)
                    #   Create best ball starting lineup
                    starters_bb = pd.DataFrame(starters_bb, columns=roster_full.columns)
                    #   Create best ball bench
                    bench_bb = roster_full[~roster_full['sleeper_id'].isin(starters_bb['sleeper_id'])]
                    bench_bb = bench_bb.sort_values(by=['POS', 'Pts'], ascending=[True, False])
                    bench_bb = bench_bb.reset_index(drop=True)
                    #   Break starters out by position
                    starters_actual_flex_loc = []
                    starters_actual_sflex_loc = []
                    starters_actual_qb_loc = starters_actual[starters_actual['POS'].str.contains('QB')].index.tolist()
                    starters_actual_rb_loc = starters_actual[starters_actual['POS'].str.contains('RB')].index.tolist()
                    starters_actual_wr_loc = starters_actual[starters_actual['POS'].str.contains('WR')].index.tolist()
                    starters_actual_te_loc = starters_actual[starters_actual['POS'].str.contains('TE')].index.tolist()
                    starters_actual_k_loc = starters_actual[starters_actual['POS'].str.contains('K')].index.tolist()
                    starters_actual_dst_loc = starters_actual[starters_actual['POS'].str.contains('DST')].index.tolist()
                    #   Fucking Taysom Hill
                    #   Players with multiple position eligibility need to be figured which position they were started in
                    Taysom_hill_id = '4381'
                    if any(starters_actual['sleeper_id'].str.contains(Taysom_hill_id)):
                        if len(starters_actual_qb_loc) > (cur_league.Num_qb + cur_league.Num_sflex):
                            TH_loc = starters_actual[starters_actual['POS'].str.contains('QB')]['sleeper_id'].values.tolist().index(Taysom_hill_id)
                            starters_actual_qb_loc.remove(starters_actual_qb_loc[TH_loc])
                        else:
                            TH_loc = starters_actual[starters_actual['POS'].str.contains('TE')]['sleeper_id'].values.tolist().index(Taysom_hill_id)
                            starters_actual_te_loc.remove(starters_actual_te_loc[TH_loc])
                    if len(starters_actual_qb_loc) > cur_league.Num_qb:
                        starters_actual_sflex_loc = starters_actual_sflex_loc + starters_actual_qb_loc[cur_league.Num_qb:]
                        starters_actual_qb_loc = starters_actual_qb_loc[:cur_league.Num_qb]
                    if len(starters_actual_rb_loc) > cur_league.Num_rb:
                        starters_actual_flex_loc = starters_actual_rb_loc[cur_league.Num_rb:]
                        starters_actual_rb_loc = starters_actual_rb_loc[:cur_league.Num_rb]
                    if len(starters_actual_wr_loc) > cur_league.Num_wr:
                        starters_actual_flex_loc = starters_actual_flex_loc + starters_actual_wr_loc[cur_league.Num_wr:]
                        starters_actual_wr_loc = starters_actual_wr_loc[:cur_league.Num_wr]
                    if len(starters_actual_te_loc) > cur_league.Num_te:
                        starters_actual_flex_loc = starters_actual_flex_loc + starters_actual_te_loc[cur_league.Num_te:]
                        starters_actual_te_loc = starters_actual_te_loc[:cur_league.Num_te]
                    if len(starters_actual_flex_loc) > cur_league.Num_flex:
                        starters_actual_sflex_loc = starters_actual_sflex_loc + starters_actual_flex_loc[cur_league.Num_flex:]
                        starters_actual_flex_loc = starters_actual_flex_loc[:cur_league.Num_flex]
                    starters_actual = starters_actual.loc[starters_actual_qb_loc + starters_actual_rb_loc + starters_actual_wr_loc + starters_actual_te_loc + starters_actual_flex_loc + starters_actual_sflex_loc + starters_actual_k_loc + starters_actual_dst_loc]
                    #   Fucking Henry Ruggs and empty roster spots
                    #   Starters who don't have a team need to be delt with and starting roster labeled as illegal
                    while len(starters_actual) < cur_league.Num_starter:
                        illegal = True
                        temp_list = starters_actual.values.tolist()
                        temp_list.append(['ILLEGAL', 0, 'ILLEGAL', 'ILLEGAL', 'NOT FOUND'])
                        starters_actual = pd.DataFrame(temp_list, columns=starters_actual.columns)
                    while len(bench_actual) > len(bench_bb):
                        illegal = True
                        temp_list = bench_bb.values.tolist()
                        temp_list.append(['ILLEGAL', 0, 'ILLEGAL', 'ILLEGAL', 'NOT FOUND'])
                        bench_bb = pd.DataFrame(temp_list, columns=bench_actual.columns)
                    starters_actual = starters_actual.reset_index(drop=True)
                    #   Start populating players weekly sheet
                    print('{0} {1} Week {2}: {3}'.format(cur_league.Name, year, week, user_name))
                    try:
                        sheet = gsheet.worksheet(user_name)
                    #   Create sheet if owner sheet doesn't already exist
                    except:
                        sheet = gsheet.add_worksheet(title=user_name, rows=(team_size+3) * (1+year_td-cur_league.Start_year) + 1*(year_td-cur_league.Start_year), cols=max_col)
                        rules = get_conditional_format_rules(sheet)
                        #   Format always bold rows & columns
                        for season in range(0, (year_td-cur_league.Start_year)+1):
                            or1 = 1 + season*(1+header_size+summary_size+cur_league.Roster_size)
                            or2 = or1+1
                            or3 = or2+1
                            or4 = or2+cur_league.Num_starter
                            or5 = or4+1
                            or6 = or5+1
                            or7 = or5+cur_league.Roster_size-cur_league.Num_starter
                            #   Header
                            format_cell_range(sheet, '{0}:{1}'.format(r1, r2), txt_fmt_bld_ctr)
                            time.sleep(sl_time)
                            #   Summary
                            format_cell_range(sheet, '{0}'.format(r5), txt_fmt_bld_ctr)
                            time.sleep(sl_time)
                            for wk in range(0, end_week):
                                #   Position Rows
                                c1 = 1 + wk*(data_num_cols+1)
                                c2 = c1 + 2
                                c3 = c2 + 1
                                c4 = c3 + 1
                                c5 = c4 + 2
                                c1 = int_to_column(c1)
                                c2 = int_to_column(c2)
                                c3 = int_to_column(c3)
                                c4 = int_to_column(c4)
                                c5 = int_to_column(c5)
                                illegal_fmt = ConditionalFormatRule(
                                    ranges=[GridRange.from_a1_range(
                                        '{0}{1}'.format(c3, or1),
                                        sheet)],
                                    booleanRule=BooleanRule(
                                        condition=BooleanCondition(
                                            type='TEXT_EQ',
                                            values=[ConditionValue(userEnteredValue='ILLEGAL')]
                                        ),
                                        format=CellFormat(backgroundColor=red)))
                                rules.append(illegal_fmt)
                                efficient_fmt = ConditionalFormatRule(
                                    ranges=[GridRange.from_a1_range(
                                        '{0}{1}'.format(c5, or1),
                                        sheet)],
                                    gradientRule=GradientRule(
                                        minpoint=InterpolationPoint(
                                            color=red,
                                            value='0.5',
                                            type='NUMBER'),
                                        midpoint=InterpolationPoint(
                                            color=yellow,
                                            value='0.75',
                                            type='NUMBER'),
                                        maxpoint=InterpolationPoint(
                                            color=green,
                                            value='1',
                                            type='NUMBER')))
                                rules.append(efficient_fmt)
                                format_cell_range(sheet, '{0}:{0}'.format(c3), txt_fmt_bld_ctr)
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}:{2}{3}'.format(c1, or1, c5, or7), Border_format('none'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}'.format(c5, or1), CellFormat(
                                    numberFormat=NumberFormat(
                                        type='PERCENT',
                                        pattern='#%')))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}:{0}{2}'.format(c2, or3, or7), CellFormat(
                                    numberFormat=NumberFormat(
                                        type='NUMBER',
                                        pattern='0.00#')))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}:{0}{2}'.format(c4, or3, or7), CellFormat(
                                    numberFormat=NumberFormat(
                                        type='NUMBER',
                                        pattern='0.00#')))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}:{2}{1}'.format(c1, or1, c5), Border_format('t'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}:{2}{1}'.format(c1, or2, c5), Border_format('b'))
                                format_cell_range(sheet, '{0}{1}:{2}{1}'.format(c1, or7, c5), Border_format('b'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}:{0}{2}'.format(c1, or1, or7), Border_format('l'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}:{0}{2}'.format(c5, or1, or7), Border_format('r'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}:{2}{1}'.format(c1, or5, c5), Border_format('tb'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}:{0}{2}'.format(c3, or3, or7), Border_format('lr'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}'.format(c1, or1), Border_format('tl'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}'.format(c5, or1), Border_format('tr'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}'.format(c1, or2), Border_format('bl'))
                                format_cell_range(sheet, '{0}{1}'.format(c1, or7), Border_format('bl'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}'.format(c5, or2), Border_format('br'))
                                format_cell_range(sheet, '{0}{1}'.format(c5, or7), Border_format('br'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}'.format(c1, or5), Border_format('tbl'))
                                format_cell_range(sheet, '{0}{1}'.format(c5, or5), Border_format('tbr'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}'.format(c3, or7), Border_format('blr'))
                                time.sleep(sl_time)
                                format_cell_range(sheet, '{0}{1}'.format(c3, or5), Border_format('all'))
                                time.sleep(sl_time)
                        rules.save()
                    # dumdum = sum(~starters_actual['sleeper_id'].isin(starters_bb['sleeper_id']))/len(starters_bb)
                    #   Position Rows
                    c1 = 1 + (week-1)*(data_num_cols+1)
                    c2 = c1 + 2
                    c3 = c2 + 1
                    c4 = c3 + 1
                    c5 = c4 + 2
                    c1 = int_to_column(c1)
                    c2 = int_to_column(c2)
                    c3 = int_to_column(c3)
                    c4 = int_to_column(c4)
                    c5 = int_to_column(c5)
                    sheet_headers = ['{0}'.format(year), 'Week {0}'.format(week), 'Actual', '', 'Ideal', 'Efficiency:', '=ROUND({0}{1}/{2}{1},2)'.format(c2, r5, c4)]
                    if illegal:
                        sheet_headers[3] = 'ILLEGAL'
                    final_list = [sheet_headers]
                    final_list.append(['Player POS', 'Name', 'Pts.', 'POS', 'Pts.', 'Name', 'Player POS'])
                    starters_actual_columns = starters_actual.columns.values
                    starters_bb_columns = starters_bb.columns.values
                    pos_names = cur_league.Starting_pos
                    starters_list = []
                    for row in range(0, cur_league.Num_starter):
                        final_list.append([starters_actual['POS'][row], starters_actual['Name'][row], starters_actual['Pts'][row],
                            pos_names[row],
                            starters_bb['Pts'][row], starters_bb['Name'][row], starters_bb['POS'][row]])
                    #final_list.append(['', 'Actual Total', round(starters_actual['Pts'].sum(),2), 'Total', round(starters_bb['Pts'].sum(),2), 'Best Possible Total', ''])
                    final_list.append(['', 'Actual Total', '=ROUND(SUM({0}{1}:{0}{2}),2)'.format(c2, r3, r4), 'Total', '=ROUND(SUM({0}{1}:{0}{2}),2)'.format(c4, r3, r4), 'Best Possible Total', ''])
                    bench_list = []
                    for row in range(0, cur_league.Roster_size-cur_league.Num_starter):
                        try:
                            final_list.append([bench_actual['POS'][row], bench_actual['Name'][row], bench_actual['Pts'][row],
                                'BN',
                                bench_bb['Pts'][row], bench_bb['Name'][row], bench_bb['POS'][row]])
                        except:
                            final_list.append(['N/A', 'Empty', 0,
                                'BN',
                                0, 'Empty', 'N/A'])
                    try:
                        sheet.update(final_list, '{0}{1}:{2}{3}'.format(c1, r1, c5, r7), raw=False)
                    except APIError:
                        sheet.add_rows(len(final_list))
                        sheet.update(final_list, '{0}{1}:{2}{3}'.format(c1, r1, c5, r7), raw=False)
                    time.sleep(sl_time)
                    if not point_count:
                        sheet.update([[0]], '{0}{1}'.format(c4, r5), raw=False)
            #   Create a year summary tab if it doesn't already exist
            try:
                sheet = gsheet.worksheet(str(year))
            except:
                r1 = 1
                r2 = r1+1
                r3 = r1+num_owners
                r4 = r3+2
                r5 = r4+1
                r6 = r4+non_playoff
                r7 = r4+num_playoff
                r8 = r4+num_owners
                weekly_table = [['Owner', 'Team', 'Max PF', 'Week 1', 'Week 2', 'Week 3', 'Week 4', 'Week 5', 'Week 6', 'Week 7', 'Week 8', 'Week 9', 'Week 10', 'Week 11', 'Week 12', 'Week 13', 'Week 14', 'Week 15', 'Week 16', 'Week 17', 'Week 18']]
                draft_table = [['Draft Order', '=B1', '=J{0}'.format(r4), '=K{0}'.format(r4), '=H{0}'.format(r4), '=N{0}'.format(r4)]]
                standings_table = [['Standings', '=B1', 'W', 'L', 'PF', 'Max PF', 'Max PF Place']]
                max_table = [['=N{0}'.format(r4), '=B1', '=M{0}'.format(r4)]]
                for owner_num in range(0,num_owners):
                    #   Weekly Max PF table
                    owner = cur_league.Users[owner_num][1]
                    r_week_score_yr = 1 + (year-cur_league.Start_year)*(1+header_size+summary_size+cur_league.Roster_size) + 2 + cur_league.Num_starter
                    wk1 = '={0}!{1}{2}'.format(owner, int_to_column((1-1)*8+5), r_week_score_yr)
                    wk2 = '={0}!{1}{2}'.format(owner, int_to_column((2-1)*8+5), r_week_score_yr)
                    wk3 = '={0}!{1}{2}'.format(owner, int_to_column((3-1)*8+5), r_week_score_yr)
                    wk4 = '={0}!{1}{2}'.format(owner, int_to_column((4-1)*8+5), r_week_score_yr)
                    wk5 = '={0}!{1}{2}'.format(owner, int_to_column((5-1)*8+5), r_week_score_yr)
                    wk6 = '={0}!{1}{2}'.format(owner, int_to_column((6-1)*8+5), r_week_score_yr)
                    wk7 = '={0}!{1}{2}'.format(owner, int_to_column((7-1)*8+5), r_week_score_yr)
                    wk8 = '={0}!{1}{2}'.format(owner, int_to_column((8-1)*8+5), r_week_score_yr)
                    wk9 = '={0}!{1}{2}'.format(owner, int_to_column((9-1)*8+5), r_week_score_yr)
                    wk10 = '={0}!{1}{2}'.format(owner, int_to_column((10-1)*8+5), r_week_score_yr)
                    wk11 = '={0}!{1}{2}'.format(owner, int_to_column((11-1)*8+5), r_week_score_yr)
                    wk12 = '={0}!{1}{2}'.format(owner, int_to_column((12-1)*8+5), r_week_score_yr)
                    wk13 = '={0}!{1}{2}'.format(owner, int_to_column((13-1)*8+5), r_week_score_yr)
                    wk14 = '={0}!{1}{2}'.format(owner, int_to_column((14-1)*8+5), r_week_score_yr)
                    wk15 = '={0}!{1}{2}'.format(owner, int_to_column((15-1)*8+5), r_week_score_yr)
                    wk16 = '={0}!{1}{2}'.format(owner, int_to_column((16-1)*8+5), r_week_score_yr)
                    wk17 = '={0}!{1}{2}'.format(owner, int_to_column((17-1)*8+5), r_week_score_yr)
                    wk18 = '={0}!{1}{2}'.format(owner, int_to_column((18-1)*8+5), r_week_score_yr)
                    weekly_table.append([owner, '', '=ROUND(SUM({0}{1}:{2}{1}),2)'.format(int_to_column(4),r2+owner_num,int_to_column(3+18)),
                        wk1, wk2, wk3, wk4, wk5, wk6, wk7, wk8, wk9, wk10, wk11, wk12, wk13, wk14, wk15, wk16, wk17, wk18])
                    #   Draft Table
                    if owner_num < num_owners-cur_league.Num_playoff_teams:
                        draft_table.append([owner_num+1, '=XLOOKUP(INDEX(SORT($N${0}:$N${1},1,FALSE()),$A{2}),$N${0}:$N${1},$I${0}:$K${1})'.format(r7+1, r8, r5+owner_num), '', '', '=XLOOKUP($B{0},$I${1}:$I${2},$H${1}:$H${2})'.format(r5+owner_num, r5, r8), '=XLOOKUP($B{0},$I${1}:$I${2},$N${1}:$N${2})'.format(r5+owner_num, r5, r8)])
                    else:
                        draft_table.append([owner_num+1, '=XLOOKUP(INDEX(SORT($H${0}:$H${1},1,FALSE()),$A{2}-$A${3}),$H${0}:$H${1},$I${0}:$K${1})'.format(r5, r7, r5+owner_num, r6), '', '', '=XLOOKUP($B{0},$I${1}:$I${2},$H${1}:$H${2})'.format(r5+owner_num, r5, r8), '=XLOOKUP($B{0},$I${1}:$I${2},$N${1}:$N${2})'.format(r5+owner_num, r5, r8)])
                    #   Season Max PF Table
                    if owner_num == 0:
                        max_table.append([owner_num+1,'=SORT($B$2:$C${0},2,FALSE)'.format(1+num_owners),''])
                    else:
                        max_table.append([owner_num+1,'',''])
                    #   Standings Table
                    standings_table.append([owner_num+1, '', '', '', '', '=XLOOKUP($I{0},$Q${1}:$Q${2},$R${1}:$R${2})'.format(r5+owner_num, r5, r8), '=INDEX(MATCH($I{0},$Q${1}:$Q${2},0))'.format(r5+owner_num, r5, r8)])
                    #   Create year summary sheet
                sheet = gsheet.add_worksheet(title=str(year), rows=(r8), cols=3+18)
                #   Write to the sheet
                sheet.update(weekly_table, 'A1:U{0}'.format(r3), raw=False)
                sheet.update(draft_table, 'A{0}:F{1}'.format(r4, r8), raw=False)
                sheet.update(standings_table, 'H{0}:N{1}'.format(r4, r8), raw=False)
                sheet.update(max_table, 'P{0}:R{1}'.format(r4, r8), raw=False)
                sheet.update([['=HYPERLINK("'+git+'", "GitHub")']], 'S{0}:S{0}'.format(r4), raw=False)
                time.sleep(sl_time)
                #   Format the sheet
                format_cell_range(sheet, 'A1:U{0}'.format(r8), CellFormat(horizontalAlignment='Center'))
                format_cell_range(sheet, 'A1:U1', txt_fmt_bld)
                format_cell_range(sheet, 'A{0}:U{0}'.format(r4), txt_fmt_bld)
                time.sleep(sl_time)
                format_cell_range(sheet, 'A1:U{0}'.format(r3), Border_format('none'))
                format_cell_range(sheet, 'A{0}:F{1}'.format(r4, r8), Border_format('none'))
                format_cell_range(sheet, 'H{0}:N{1}'.format(r4, r8), Border_format('none'))
                format_cell_range(sheet, 'P{0}:R{1}'.format(r4, r8), Border_format('none'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'A{0}:U{0}'.format(r3), Border_format('b'))
                format_cell_range(sheet, 'A{0}:F{0}'.format(r6), Border_format('b'))
                format_cell_range(sheet, 'A{0}:F{0}'.format(r8), Border_format('b'))
                format_cell_range(sheet, 'H{0}:N{0}'.format(r7), Border_format('b'))
                format_cell_range(sheet, 'H{0}:N{0}'.format(r8), Border_format('b'))
                format_cell_range(sheet, 'P{0}:R{0}'.format(r8), Border_format('b'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'A2:A{0}'.format(r3), Border_format('l'))
                format_cell_range(sheet, 'A{0}:A{1}'.format(r4, r8), Border_format('l'))
                format_cell_range(sheet, 'H{0}:H{1}'.format(r4, r8), Border_format('l'))
                format_cell_range(sheet, 'P{0}:P{1}'.format(r4, r8), Border_format('l'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'U2:U{0}'.format(r3), Border_format('r'))
                format_cell_range(sheet, 'F{0}:F{1}'.format(r4, r8), Border_format('r'))
                format_cell_range(sheet, 'N{0}:N{1}'.format(r4, r8), Border_format('r'))
                format_cell_range(sheet, 'R{0}:R{1}'.format(r4, r8), Border_format('r'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'A2', Border_format('tl'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'U2', Border_format('tr'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'A1:U1', Border_format('tb'))
                format_cell_range(sheet, 'A{0}:F{0}'.format(r4), (Border_format('tb')))
                format_cell_range(sheet, 'H{0}:N{0}'.format(r4), Border_format('tb'))
                format_cell_range(sheet, 'P{0}:R{0}'.format(r4), Border_format('tb'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'A{0}'.format(r3), Border_format('bl'))
                format_cell_range(sheet, 'A{0}'.format(r6), Border_format('bl'))
                format_cell_range(sheet, 'H{0}'.format(r7), Border_format('bl'))
                format_cell_range(sheet, 'A{0}'.format(r8), Border_format('bl'))
                format_cell_range(sheet, 'H{0}'.format(r8), Border_format('bl'))
                format_cell_range(sheet, 'P{0}'.format(r8), Border_format('bl'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'U{0}'.format(r3), Border_format('br'))
                format_cell_range(sheet, 'F{0}'.format(r6), Border_format('br'))
                format_cell_range(sheet, 'F{0}'.format(r8), Border_format('br'))
                format_cell_range(sheet, 'N{0}'.format(r7), Border_format('br'))
                format_cell_range(sheet, 'N{0}'.format(r8), Border_format('br'))
                format_cell_range(sheet, 'R{0}'.format(r8), Border_format('br'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'C2:C{0}'.format(r3), Border_format('lr'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'A{0}'.format(r4), Border_format('tbl'))
                format_cell_range(sheet, 'H{0}'.format(r4), Border_format('tbl'))
                format_cell_range(sheet, 'P{0}'.format(r4), Border_format('tbl'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'F{0}'.format(r4), Border_format('tbr'))
                format_cell_range(sheet, 'N{0}'.format(r4), Border_format('tbr'))
                format_cell_range(sheet, 'R{0}'.format(r4), Border_format('tbr'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'C{0}'.format(r3), Border_format('blr'))
                time.sleep(sl_time)
                format_cell_range(sheet, 'C1:C1', Border_format('all'))
                time.sleep(sl_time)
                #   Format all the color gradient for week
                rules = get_conditional_format_rules(sheet)
                for wk in range(0, end_week+1):
                    c = int_to_column(3 + wk)
                    efficient_fmt = ConditionalFormatRule(
                        ranges=[GridRange.from_a1_range('{0}{1}:{0}{2}'.format(c, r2, r3), sheet)],
                        gradientRule=GradientRule(
                            minpoint=InterpolationPoint(
                                color=red,
                                type='MIN'),
                            midpoint=InterpolationPoint(
                                color=white,
                                type='PERCENT',
                                value='50'),
                            maxpoint=InterpolationPoint(
                                color=green,
                                type='MAX')))
                    rules.append(efficient_fmt)
                rules.save()
            #   Update year summary with the latest standings
            r1 = 1
            r2 = r1+1
            r3 = r1+num_owners
            r4 = r3+2
            r5 = r4+1
            r6 = r4+non_playoff
            r7 = r4+num_playoff
            r8 = r4+num_owners
            cur_league.Update(end_week, year)
            sheet = gsheet.worksheet(str(year))
            sheet.update(cur_league.Standings, 'I{0}:L{1}'.format(r5, r8), raw=False)
            time.sleep(sl_time)
            #   Playoffs
            playoffs = cur_league.League.get_playoff_winners_bracket()
            rd1 = []
            rd2 = []
            rd3 = []
            for game in range(0, len(playoffs)):
                match playoffs[game]['r']:
                    case 1:
                        rd1 = rd1 + [game]
                    case 2:
                        try:
                            playoffs[game]['t2_from']['w']
                            rd2 = rd2 + [game]
                        except:
                            pass
                    case 3:
                        try:
                            if ('w' in playoffs[game]['t2_from'] and 'w' in playoffs[game]['t1_from']) or ('l' in playoffs[game]['t2_from'] and 'l' in playoffs[game]['t1_from']):
                                rd3 = rd3 + [game]
                        except:
                            pass
            elim1 = []
            elim2 = []
            elim3 = []
            temp_all = []
            for game in rd1:
                try:
                    elim1 = elim1 + [playoffs[game]['l']-1]
                except:
                    continue
            for game in rd2:
                try:
                    elim2 = elim2 + [playoffs[game]['l']-1]
                except:
                    continue
            for game in rd3:
                try:
                    elim3 = elim3 + [playoffs[game]['l']-1]
                except:
                    continue
            owner_team = sheet.get('A{0}:B{1}'.format(r2, r3))
            bb_standings = sheet.get('Q{0}:Q{1}'.format(r5, r8))
            playoff_teams = []
            for team in range(0, num_playoff):
                temp_team = cur_league.Standings[team][0]
                for all_team in range(0, len(owner_team)):
                    if owner_team[all_team][0]==temp_team:
                    #if owner_team[all_team][1]==temp_team:
                        temp_num = all_team
                        for bb_place in range(0, len(bb_standings)):
                            if temp_team==bb_standings[bb_place][0]:
                                temp_bb = bb_place
                                playoff_teams.append({'team': temp_team, 'num': temp_num, 'bb': temp_bb})
                                continue
                        continue
            try:
                fourth = list(set(elim2).intersection(set(elim3)))[0]
                second = list(set([fourth])^(set(elim3)))[0]
                third = list(set([fourth])^(set(elim2)))[0]
                for team in playoff_teams:
                    if team['num'] not in elim1 and team['num'] not in elim2 and team['num'] not in elim3:
                        first = team
                        continue
                rd1_elim = []
                for team in playoff_teams:
                    if team['num'] in elim1:
                        rd1_elim.append(team)
                rd1_elim = sorted(rd1_elim, key=lambda x: x['bb'], reverse=True)
                playoff_draft = []
                for team in rd1_elim:
                    playoff_draft.append(team['team'])
                for team in playoff_teams:
                    if team['num'] == fourth:
                        playoff_draft.append(team['team'])
                        continue
                for team in playoff_teams:
                    if team['num'] == third:
                        playoff_draft.append(team['team'])
                        continue
                for team in playoff_teams:
                    if team['num'] == second:
                        playoff_draft.append(team['team'])
                        continue
                playoff_draft.append(first['team'])
                po_draft_tbl = []
                for row in range(r6+1, r8+1):
                    po_draft_tbl.append([playoff_draft[row-(r6+1)], '=XLOOKUP(B{0},I{1}:I{2},J{1}:J{2})'.format(row, r5, r7), '=XLOOKUP(B{0},I{1}:I{2},K{1}:K{2})'.format(row, r5, r7)])
                sheet.update(po_draft_tbl, 'B{0}:D{1}'.format(r6+1, r8), raw=False)
            except:
                continue