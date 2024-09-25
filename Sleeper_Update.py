import pandas as pd
import gspread
import time
from sleeper_wrapper import League, Players
import nfl_data_py as nfl
import openpyxl as pyxl

class Dumpster_Dynasty:

    def __init__(self, *args):
        self.args = args
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
        self.Sheet_url = 'https://docs.google.com/spreadsheets/d/1ko1XnttApOFkA1x1RF95Ls5C-aM98p766ht058Il350/edit?pli=1&gid=962080300#gid=962080300'
        self.Service_account = r'C:\Users\Jed\AppData\Local\Programs\Python\Python310\Lib\site-packages\gspread\dumpster-dynasty-bb.json'
        self.Credentials = {"installed":{"client_id":"215364173021-j261liv1nscj8tj4es9do2mun9t4a603.apps.googleusercontent.com","project_id":"dumpster-dynasty-bb","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
        self.Starting_pos = starting_pos
        self.League = League(1128506799759429632)

    def Update(self, week, season, *args):
        self.Rosters = self.League.get_rosters()
        all_users = self.League.get_users()
        self.Matchups = self.League.get_matchups(week)
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
        # 
        # self.Scoreboards = self.League.get_scoreboards(self.Rosters, self.Matchups, self.Users, "pts_half_ppr", season, week)

        # self.score1 = self.Scoreboards[1][0]
        # self.score2 = self.Scoreboards[1][1]
        # self.score3 = self.Scoreboards[2][0]
        # self.score4 = self.Scoreboards[2][1]
        # self.score5 = self.Scoreboards[3][0]
        # self.score6 = self.Scoreboards[3][1]
        # self.score7 = self.Scoreboards[4][0]
        # self.score8 = self.Scoreboards[4][1]
        # self.score9 = self.Scoreboards[5][0]
        # self.score10 = self.Scoreboards[5][1]

class RIP_Harambe:

    def __init__(self, *args):
        self.args = args
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
        self.Sheet_url = 'https://docs.google.com/spreadsheets/d/1uy91bw37lcMzz2gcw7UDDgI6NJmkR887EEvmDsJHfjc/edit?gid=0#gid=0'
        self.Service_account = r'C:\Users\Jed\AppData\Local\Programs\Python\Python310\Lib\site-packages\gspread\dumpster-dynasty-bb.json'
        self.Credentials = {"installed":{"client_id":"215364173021-j261liv1nscj8tj4es9do2mun9t4a603.apps.googleusercontent.com","project_id":"dumpster-dynasty-bb","auth_uri":"https://accounts.google.com/o/oauth2/auth","token_uri":"https://oauth2.googleapis.com/token","auth_provider_x509_cert_url":"https://www.googleapis.com/oauth2/v1/certs","client_secret":"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","redirect_uris":["http://localhost"]}}
        self.Starting_pos = starting_pos
        self.League = League(1088873848449097728)

    def Update(self, week, season, *args):
        self.Rosters = self.League.get_rosters()
        all_users = self.League.get_users()
        self.Matchups = self.League.get_matchups(week)
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
        # 
        # self.Scoreboards = self.League.get_scoreboards(self.Rosters, self.Matchups, self.Users, "pts_half_ppr", season, week)

        # self.score1 = self.Scoreboards[1][0]
        # self.score2 = self.Scoreboards[1][1]
        # self.score3 = self.Scoreboards[2][0]
        # self.score4 = self.Scoreboards[2][1]
        # self.score5 = self.Scoreboards[3][0]
        # self.score6 = self.Scoreboards[3][1]
        # self.score7 = self.Scoreboards[4][0]
        # self.score8 = self.Scoreboards[4][1]
        # self.score9 = self.Scoreboards[5][0]
        # self.score10 = self.Scoreboards[5][1]

def int_to_column(n):
    result = ''
    while n > 0:
        n, remainder = divmod(n-1, 26)
        result = chr(65+remainder) + result
    return result

def Player_stat_score(stats, League, *args):
    match League:
        case 'Dumpster':
            pass_yd = 0.04
            pass_td = 4
            pass_conv2 = 2
            pass_int = -1
            rush_yd = 0.1
            rush_td = 6
            rush_conv2 = 2
            rec = 0.5
            rec_yd = 0.1
            rec_td = 6
            rec_conv2 = 2
            fg_yd = 0.1
            pat_make = 1
            fg_miss_0_19 = -4
            fg_miss_20_29 = -3
            fg_miss_30_39 = -2
            fg_miss_40_49 = -1
            fg_miss_50_59 = -1
            pat_miss = -1
            def_td = 6
            def_allow_0 = 10
            def_allow_0_6 = 7
            def_allow_7_13 = 4
            def_allow_14_20 = 1
            def_allow_21_27 = 0
            def_allow_28_34 = -1
            def_allow_35 = -4
            sack = 1
            pick = 2
            f_recover = 2
            safe = 2
            fumble_force = 1
            block_kick = 2
            st_td = 6
            st_ff = 1
            st_fr = 1
            fumble_loss = -2
            fumble_td = 6

    pts = 0
    pts = pts + pass_yd * stats['passing_yards'].values[0]
    pts = pts + pass_td * stats['passing_tds'].values[0]
    pts = pts + pass_conv2 * stats['passing_2pt_conversions'].values[0]
    pts = pts + pass_int * stats['interceptions'].values[0]
    pts = pts + rush_yd * stats['rushing_yards'].values[0]
    pts = pts + rush_td * stats['rushing_tds'].values[0]
    pts = pts + rush_conv2 * stats['rushing_2pt_conversions'].values[0]
    pts = pts + rec * stats['receptions'].values[0]
    pts = pts + rec_yd * stats['receiving_yards'].values[0]
    pts = pts + rec_td * stats['receiving_tds'].values[0]
    pts = pts + rec_conv2 * stats['receiving_2pt_conversions'].values[0]
    pts = pts + st_td * stats['special_teams_tds'].values[0]
    pts = pts + fumble_loss * (stats['sack_fumbles_lost'].values[0] + stats['rushing_fumbles_lost'].values[0] + stats['receiving_fumbles_lost'].values[0])

    pts = round(pts, 2)
    
    return pts

def Best_player(possible, starters, *args):
    if starters:
        starters_df = pd.DataFrame(starters, columns=possible.columns)
        possible = possible[~possible['gsis_id'].isin(starters_df['gsis_id'])]
    starters.append(possible.loc[possible.idxmax().Pts].values)

    return starters

if __name__ == "__main__":
    start_week = 3
    end_week = 3
    year = 2024
    DST_list = ['ARI', 'ATL', 'BAL', 'BUF', 'CAR', 'CHI', 'CIN', 'CLE', 'DAL', 'DEN', 'DET', 'GB', 'HOU', 'IND', 'JAX', 'KC', 'LAC', 'LAR', 'LV', 'MIA', 'MIN', 'NE', 'NYG', 'NYJ', 'NO', 'PHI', 'PIT', 'SEA', 'SF', 'TB', 'TEN', 'WAS']
    # gc, authorized_user = gspread.oauth_from_dict(credentials)
    Dumpster = Dumpster_Dynasty()
    Harambe = RIP_Harambe()
    leagues_all = [Harambe, Dumpster]
    player_table = pd.read_csv('https://github.com/nflverse/nflverse-data/releases/download/rosters/roster_2024.csv')
    # all_players = Players().get_all_players('nfl')
    # week_data = nfl.import_weekly_data([year])
    for cur_league in leagues_all:
        gc = gspread.service_account(filename=cur_league.Service_account)
        gsheet = gc.open_by_url(cur_league.Sheet_url)
        for week in range(start_week, end_week+1):
            cur_league.Update(week, year)
            sheet = gsheet.worksheet(str(year))
            sheet.update(cur_league.Standings, 'I{0}:L{1}'.format(1+1+len(cur_league.Users)+2, 1+1+len(cur_league.Users)+2+len(cur_league.Users)))
            for owner in range(0, len(cur_league.Users)):
                illegal = False
                roster = cur_league.Matchups[owner]['players']
                starters = cur_league.Matchups[owner]['starters']
                bench = list(set(roster) ^ set(starters))
                user_id = cur_league.Users[owner][0]
                user_name = cur_league.Users[owner][1]
                roster_names = []
                roster_pos = []
                roster_data = []
                bench_pts = []
                starter_pts = []
                roster_data_dict = cur_league.Matchups[owner]['players_points']
                for sleeper_id, pts in roster_data_dict.items():
                    if sleeper_id in starters:
                        starter_pts.append([sleeper_id, pts])
                    else:
                        bench_pts.append([sleeper_id, pts])
                    roster_data.append([sleeper_id, pts])
                roster_gsis_ids = []
                roster_sleeper_ids = []
                for player in range(len(roster)):
                    if str(roster[player]) in DST_list:
                        player_name = roster[player]
                        player_gsis_id = roster[player]
                        player_pos = 'DST'
                    else:
                        player_nfl = player_table[player_table['sleeper_id'].isin([int(roster[player])])][['full_name', 'gsis_id', 'position', 'depth_chart_position']]
                        try:
                            player_pos = player_nfl['position'].values[0]
                            player_name = player_nfl['full_name'].values[0]
                            player_gsis_id = player_nfl['gsis_id'].values[0]
                        except:
                            player_pos = 'NOT FOUND'
                            player_name = 'NOT FOUND'
                            player_gsis_id = 'NOT FOUND'
                        if player_name == 'Taysom Hill':
                            player_pos = 'TE & QB'
                    roster_data[player].append(player_name)
                    roster_data[player].append(player_pos)
                    roster_data[player].append(player_gsis_id)
                roster_full = pd.DataFrame(roster_data, columns=['sleeper_id', 'Pts', 'Name', 'POS', 'gsis_id'])
                starters_actual = roster_full[roster_full['sleeper_id'].isin(starters)]
                bench_actual = roster_full[~roster_full['sleeper_id'].isin(starters)]
                roster_full = roster_full.sort_values(by=['POS', 'Pts'], ascending=[True, False])
                starters_actual = starters_actual.sort_values(by=['POS', 'Pts'], ascending=[True, False])
                bench_actual = bench_actual.sort_values(by=['POS', 'Pts'], ascending=[True, False])
                bench_actual = bench_actual.reset_index(drop=True)
                roster_qb = roster_full[roster_full['POS'].str.contains('QB')]
                roster_rb = roster_full[roster_full['POS'].str.contains('RB')]
                roster_wr = roster_full[roster_full['POS'].str.contains('WR')]
                roster_te = roster_full[roster_full['POS'].str.contains('TE')]
                roster_k = roster_full[roster_full['POS'].str.contains('K')]
                roster_dst = roster_full[roster_full['POS'].str.contains('DST')]
                roster_flex = roster_full[roster_full['POS'].str.contains('WR')|roster_full['POS'].str.contains('RB')|roster_full['POS'].str.contains('TE')]
                roster_sflex = roster_full[roster_full['POS'].str.contains('QB')|roster_full['POS'].str.contains('WR')|roster_full['POS'].str.contains('RB')|roster_full['POS'].str.contains('TE')]
                starters_bb = []
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
                starters_bb = pd.DataFrame(starters_bb, columns=roster_full.columns)
                bench_bb = roster_full[~roster_full['sleeper_id'].isin(starters_bb['sleeper_id'])]
                bench_bb = bench_bb.sort_values(by=['POS', 'Pts'], ascending=[True, False])
                bench_bb = bench_bb.reset_index(drop=True)
                starters_actual_flex_loc = []
                starters_actual_sflex_loc = []
                starters_actual_qb_loc = starters_actual[starters_actual['POS'].str.contains('QB')].index.tolist()
                starters_actual_rb_loc = starters_actual[starters_actual['POS'].str.contains('RB')].index.tolist()
                starters_actual_wr_loc = starters_actual[starters_actual['POS'].str.contains('WR')].index.tolist()
                starters_actual_te_loc = starters_actual[starters_actual['POS'].str.contains('TE')].index.tolist()
                starters_actual_k_loc = starters_actual[starters_actual['POS'].str.contains('K')].index.tolist()
                starters_actual_dst_loc = starters_actual[starters_actual['POS'].str.contains('DST')].index.tolist()
                #   Fucking Taysom Hill
                Taysom_hill_id = '4381'
                if any(starters_actual['sleeper_id'].str.contains(Taysom_hill_id)):
                    if len(starters_actual_qb_loc) > (cur_league.Num_qb + cur_league.Num_sflex):
                        starters_actual_qb_loc = starters_actual_qb_loc[~starters_actual_qb_loc.isin([Taysom_hill_id])]
                    else:
                        starters_actual_te_loc = starters_actual_te_loc[~starters_actual_te_loc.isin([Taysom_hill_id])]
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
                #   Fucking Henry Ruggs
                if len(starters_actual) < len(starters_bb):
                    illegal = True
                    temp_list = starters_actual.values.tolist()
                    temp_list.append(['ILLEGAL', 0, 'ILLEGAL', 'ILLEGAL', 'NOT FOUND'])
                    starters_actual = pd.DataFrame(temp_list, columns=starters_actual.columns)
                starters_actual = starters_actual.reset_index(drop=True)
                data_num_cols = 7
                week_start_col = (week-1) * data_num_cols + week
                print(user_name)
                try:
                    sheet = gsheet.worksheet(user_name)
                except:
                    sheet = gsheet.add_worksheet(title=user_name, rows=50, cols=100)
                    sheet.format(['A1:2'], {
                        'textFormat': {
                            'bold': True
                        },
                        'horizontalAlignment': 'Center'
                    })
                    sheet.format(['A14:CV14'], {
                        'textFormat': {'bold': True
                        },
                        'horizontalAlignment': 'Center'
                    })
                    sheet.format(['D1:D50'], {
                        'textFormat': {'bold': True
                        },
                        'horizontalAlignment': 'Center'
                    })
                dumdum = sum(~starters_actual['sleeper_id'].isin(starters_bb['sleeper_id']))/len(starters_bb)
                sheet_col_start = int_to_column(week+((week-1)*7))
                sheet_col_end = int_to_column(week+((week-1)*7)+6)
                sheet_headers = ['{0}'.format(year), 'Week {0}'.format(week), 'Actual', '', 'Ideal', 'DumDum Rate:', dumdum]
                if illegal:
                    sheet_headers[3] = 'ILLEGAL ROSTER'
                final_list = [sheet_headers]
                final_list.append(['Player POS', 'Name', 'Pts.', 'Roster POS', 'Pts.', 'Name', 'Player POS'])
                starters_actual_columns = starters_actual.columns.values
                starters_bb_columns = starters_bb.columns.values
                pos_names = cur_league.Starting_pos
                starters_list = []
                for row in range(0, len(starters_bb)):
                    final_list.append([starters_actual['POS'][row], starters_actual['Name'][row], starters_actual['Pts'][row],
                                          pos_names[row],
                                          starters_bb['Pts'][row], starters_bb['Name'][row], starters_bb['Pts'][row]])
                final_list.append(['', 'Actual Total', round(starters_actual['Pts'].sum(),2), 'Total', round(starters_bb['Pts'].sum(),2), 'Best Possible Total', ''])
                bench_list = []
                for row in range(0, len(bench_actual)):
                    final_list.append([bench_actual['POS'][row], bench_actual['Name'][row], bench_actual['Pts'][row],
                                          'BN',
                                          bench_bb['Pts'][row], bench_bb['Name'][row], bench_bb['Pts'][row]])
                sheet.update(final_list, '{0}1:{1}{2}'.format(sheet_col_start, sheet_col_end, 2+len(starters_bb)+1+len(bench_actual)))
