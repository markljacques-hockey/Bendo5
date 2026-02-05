import streamlit as st
import pandas as pd
import urllib.parse
import io
from datetime import datetime, timedelta

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="Hockey Team Balancer")

# --- 2. CACHED DATA LOADER ---
@st.cache_data
def load_excel_data(uploaded_file):
    try:
        return pd.read_excel(uploaded_file, sheet_name=None, header=1)
    except Exception as e:
        return None

# --- 3. HELPER FUNCTIONS ---
def snake_draft(players):
    players = players.reset_index(drop=True)
    team_a_list = []
    team_b_list = []
    
    for i in range(len(players)):
        player = players.iloc[i]
        if i % 4 == 0 or i % 4 == 3:
            team_a_list.append(player)
        else:
            team_b_list.append(player)
            
    cols = players.columns
    if not team_a_list: df_a = pd.DataFrame(columns=cols)
    else: df_a = pd.DataFrame(team_a_list, columns=cols)
        
    if not team_b_list: df_b = pd.DataFrame(columns=cols)
    else: df_b = pd.DataFrame(team_b_list, columns=cols)
        
    return df_a, df_b

def format_team_list(df, team_name):
    if df.empty: return f"{team_name} (0 players):\n"
    txt = f"{team_name} ({len(df)} players):\n"
    if 'Position' in df.columns and 'Full Name' in df.columns:
        df_sorted = df.sort_values(by=['Position', 'Full Name'], ascending=[True, True])
        for _, row in df_sorted.iterrows():
            txt += f"- {row['Full Name']} ({row['Position']})\n"
    return txt

def get_top_n_score(df, n):
    if df.empty or n <= 0: return 0
    return df.sort_values(by='Score', ascending=False).head(n)['Score'].sum()

def clean_name_key(name):
    """Creates a simplified string for matching (lowercase, no spaces)."""
    if pd.isna(name): return ""
    return str(name).lower().replace(" ", "").strip()

def find_col_case_insensitive(df, target_names):
    if isinstance(target_names, str): target_names = [target_names]
    target_names = [t.lower() for t in target_names]
    for col in df.columns:
        if str(col).strip().lower() in target_names:
            return col
    return None

def get_birthday_message(players_df, bday_col):
    """
    Checks for birthdays in the current calendar week (Mon-Sun) for ALL players provided.
    """
    if not bday_col or bday_col not in players_df.columns:
        return "", []
        
    today = datetime.now()
    
    # SUNDAY LOOKAHEAD LOGIC
    # If today is Sunday (weekday 6), look at NEXT week.
    if today.weekday() == 6:
        start_of_week = today + timedelta(days=1)
    else:
        start_of_week = today - timedelta(days=today.weekday())
        
    # Normalize time to start of day / end of day
    start_of_week = start_of_week.replace(hour=0, minute=0, second=0, microsecond=0)
    end_of_week = start_of_week + timedelta(days=6)
    end_of_week = end_of_week.replace(hour=23, minute=59, second=59, microsecond=999999)

    celebrants = []
    
    # Ensure date column is datetime format just in case
    # Copy to avoid SettingWithCopy warnings if slice passed
    working_df = players_df.copy()
    working_df[bday_col] = pd.to_datetime(working_df[bday_col], errors='coerce')

    for _, row in working_df.iterrows():
        bday = row[bday_col]
        if pd.isna(bday) or str(bday).strip() == '': continue
            
        try:
            if isinstance(bday, (pd.Timestamp, datetime)):
                # Check Current, Prev, and Next year to handle boundaries/New Year
                candidates = []
                years = [today.year, today.year - 1, today.year + 1]
                for y in years:
                    try:
                        candidates.append(bday.replace(year=y))
                    except ValueError:
                        # Handle leap years (Feb 29 -> Mar 1)
                        candidates.append(bday.replace(year=y, month=3, day=1))
                
                # Check match
                if any(start_of_week <= c <= end_of_week for c in candidates):
                    # Prefer "Full Name", fallback to Name or construct it
                    if 'Full Name' in row:
                        name = row['Full Name']
                    elif 'Name' in row:
                        name = row['Name']
                    else:
                        name = "Unknown Player"
                    
                    if name not in celebrants:
                        celebrants.append(name)
        except Exception: continue
            
    if not celebrants: return "", []

    # Grammar
    names_str = " and ".join([", ".join(celebrants[:-1]), celebrants[-1]] if len(celebrants) > 2 else celebrants)
    verb = "is" if len(celebrants) == 1 else "are"
    noun = "birthday" if len(celebrants) == 1 else "birthdays"
    
    msg = f"🎉 Congratulations to {names_str} who {verb} celebrating their {noun} this week!\n\n"
    return msg, celebrants

# --- 4. MAIN APP ---
st.title("🏒 Hockey Team Generator")
st.write("Upload your Excel player sheet.")

uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls', 'xlsm'])

if uploaded_file is not None:
    all_sheets = load_excel_data(uploaded_file)
    if all_sheets is None:
        st.error("Error reading file.")
        st.stop()
        
    all_sheet_names = list(all_sheets.keys())
    
    # 1. IDENTIFY MASTER SHEET (for lookup)
    master_sheet_name = next((s for s in all_sheet_names if "master" in s.lower()), None)
    
    # 2. FILTER DROPDOWN OPTIONS
    # Explicitly exclude "Master" and "Instructions" regardless of case/spacing
    block_list = ["master", "instructions", "instruction"]
    
    valid_sheets = []
    for s in all_sheet_names:
        clean_name = s.strip().lower()
        if clean_name not in block_list:
            valid_sheets.append(s)
    
    if not valid_sheets:
        st.error("No valid daily sheets found. (Hidden: Master, Instructions).")
        st.stop()
        
    selected_sheet = st.selectbox("Select the Sheet to use:", valid_sheets)
    
    try:
        df = all_sheets[selected_sheet].copy()
    except KeyError:
        st.stop()

    # --- NORMALIZE COLUMNS ---
    col_name = find_col_case_insensitive(df, ['name', 'full name', 'fullname'])
    col_first = find_col_case_insensitive(df, ['first_name', 'first name', 'firstname'])
    col_last = find_col_case_insensitive(df, ['last_name', 'last name', 'lastname'])
    col_avail = find_col_case_insensitive(df, ['availability', 'avail'])
    col_choice = find_col_case_insensitive(df, ['1st choice', '1stchoice', 'position', 'pos'])
    col_score = find_col_case_insensitive(df, ['score', 'rating', 'skill'])
    
    # Validation
    if not (col_name or (col_first and col_last)):
        st.error(f"Missing Name columns in '{selected_sheet}'. Need 'Name' OR 'First_name'/'Last_name'.")
        st.stop()
    if not col_avail or not col_score or not col_choice:
        st.error(f"Missing one of: Availability, Score, 1st Choice in '{selected_sheet}'.")
        st.stop()

    df['Availability'] = df[col_avail].astype(str).str.strip().str.upper()
    df['1st Choice'] = df[col_choice].astype(str).str.strip().str.upper()
    df['Score'] = df[col_score]
    
    if col_name:
        df['Full Name'] = df[col_name].astype(str).str.strip()
    else:
        df['Full Name'] = df[col_first].astype(str).str.strip() + ' ' + df[col_last].astype(str).str.strip()
    
    col_2nd = find_col_case_insensitive(df, ['2nd choice', '2ndchoice'])
    col_email = find_col_case_insensitive(df, ['email', 'e-mail'])
    col_reg = find_col_case_insensitive(df, ['reg/spare', 'status'])
    
    df['2nd Choice'] = df[col_2nd].fillna('').astype(str).str.strip().str.upper() if col_2nd else ''
    df['Email'] = df[col_email].fillna('').astype(str).str.strip() if col_email else ''
    df['Reg/Spare'] = df[col_reg] if col_reg else 'R'

    # --- MASTER SHEET LOOKUP (BIRTHDAYS) ---
    bday_col_name = 'B-day' 
    birthday_source_df = None # This will hold the DF we check for birthdays
    
    if master_sheet_name:
        try:
            df_master = all_sheets[master_sheet_name].copy()
            
            # Identify Cols in Master
            m_col_name = find_col_case_insensitive(df_master, ['name', 'full name'])
            m_col_first = find_col_case_insensitive(df_master, ['first_name', 'first name'])
            m_col_last = find_col_case_insensitive(df_master, ['last_name', 'last name'])
            m_col_bday = find_col_case_insensitive(df_master, ['b-day', 'bday', 'birthday', 'dob'])
            
            if (m_col_name or (m_col_first and m_col_last)) and m_col_bday:
                # 1. Prepare Master Names
                if m_col_name:
                    df_master['MatchKey'] = df_master[m_col_name].apply(clean_name_key)
                    df_master['Full Name'] = df_master[m_col_name]
                else:
                    df_master['MatchKey'] = (df_master[m_col_first].astype(str) + df_master[m_col_last].astype(str)).apply(clean_name_key)
                    df_master['Full Name'] = df_master[m_col_first].astype(str).str.strip() + ' ' + df_master[m_col_last].astype(str).str.strip()
                
                # 2. Standardize Master Birthday
                df_master[bday_col_name] = pd.to_datetime(df_master[m_col_bday], errors='coerce')
                
                # 3. SET BIRTHDAY SOURCE TO MASTER (Check everyone in Master)
                birthday_source_df = df_master
                
                # 4. Also map to Daily DF for redundancy (optional but good)
                df_master_clean = df_master.drop_duplicates(subset=['MatchKey'])
                bday_map = df_master_clean.set_index('MatchKey')[bday_col_name].to_dict()
                df['MatchKey'] = df['Full Name'].apply(clean_name_key)
                df[bday_col_name] = df['MatchKey'].map(bday_map)
                
            else:
                st.warning("Master sheet found but missing Name or B-day columns.")
                # Fallback: Use daily sheet
                birthday_source_df = df
        except Exception as e:
            st.warning(f"Error reading Master sheet: {e}")
            birthday_source_df = df
    else:
        # Fallback to daily
        daily_bday = find_col_case_insensitive(df, ['b-day', 'bday', 'birthday'])
        if daily_bday: 
            bday_col_name = daily_bday
            df[bday_col_name] = pd.to_datetime(df[bday_col_name], errors='coerce')
        birthday_source_df = df

    # --- FILTER PLAYING PLAYERS ---
    available = df[df['Availability'].str.startswith('Y')].copy()
    
    if available.empty:
        st.error(f"No players marked as 'Y' or 'Yes' in sheet '{selected_sheet}'.")
        st.stop()
    
    # --- TARGETS ---
    total_players = len(available)
    BASE_F, BASE_D, MIN_D_CRITICAL = 12, 8, 6
    if total_players <= 20:
        target_f, target_d = BASE_F, BASE_D
    else:
        extras = total_players - 20
        add_to_f = min(extras, 6); extras -= add_to_f
        add_to_d = min(extras, 4)
        target_f, target_d = BASE_F + add_to_f, BASE_D + add_to_d
    
    st.info(f"**Roster Strategy ({selected_sheet}):** Found {total_players} players. Aiming for **{target_f} Forwards** and **{target_d} Defensemen**.")
    
    # --- BIRTHDAY DEBUGGER ---
    # We now pass birthday_source_df (likely Master) to check EVERYONE
    bday_msg, bday_names = get_birthday_message(birthday_source_df, bday_col_name)
    
    with st.expander("🎂 Birthday Checker (Debug Info)", expanded=True):
        today = datetime.now()
        if today.weekday() == 6: 
            s_week = today + timedelta(days=1)
            st.write("📅 **Sunday detected:** Scanning NEXT week.")
        else:
            s_week = today - timedelta(days=today.weekday())
        e_week = s_week + timedelta(days=6)
        
        st.write(f"**Scanning Window:** {s_week.strftime('%b %d')} — {e_week.strftime('%b %d')}")
        if birthday_source_df is not None:
             st.write(f"**Checking Source:** {'Master Sheet' if master_sheet_name else 'Daily Sheet'} ({len(birthday_source_df)} players checked)")
        
        if bday_names:
            st.success(f"**Birthdays this week:** {', '.join(bday_names)}")
        else:
            st.info("No birthdays found for this week in the entire list.")

    # --- SORT & POOLS ---
    available = available.sample(frac=1).reset_index(drop=True)
    available['Status_Rank'] = available['Reg/Spare'].apply(lambda x: 0 if str(x).strip().upper() == 'R' else 1)
    available = available.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])

    pool_d = available[available['1st Choice'] == 'D'].copy()
    pool_f = available[available['1st Choice'] == 'F'].copy()

    # --- FILL GAPS ---
    if len(pool_d) < MIN_D_CRITICAL:
        needed = MIN_D_CRITICAL - len(pool_d)
        candidates = pool_f[pool_f['2nd Choice'] == 'D']
        if not candidates.empty:
            converts = candidates.head(needed)
            pool_d = pd.concat([pool_d, converts]); pool_f = pool_f.drop(converts.index)
            st.warning(f"⚠️ Critical D Shortage: Moved {len(converts)} F -> D: {', '.join(converts['Full Name'])}")
    
    if len(pool_f) < target_f:
        needed = target_f - len(pool_f)
        surplus_d = len(pool_d) - MIN_D_CRITICAL
        if surplus_d > 0:
            can_take = min(needed, surplus_d)
            candidates = pool_d[pool_d['2nd Choice'] == 'F']
            if not candidates.empty:
                converts = candidates.head(can_take)
                pool_f = pd.concat([pool_f, converts]); pool_d = pool_d.drop(converts.index)
                st.info(f"Moved {len(converts)} D -> F: {', '.join(converts['Full Name'])}")

    if len(pool_d) < target_d:
        d_shortage = target_d - len(pool_d)
        surplus_f = len(pool_f) - target_f
        if surplus_f > 0:
            amount_to_move = min(d_shortage, surplus_f)
            candidates = pool_f[pool_f['2nd Choice'] == 'D']
            if not candidates.empty:
                converts = candidates.head(amount_to_move)
                pool_d = pd.concat([pool_d, converts]); pool_f = pool_f.drop(converts.index)
                st.info(f"Moved {len(converts)} F -> D: {', '.join(converts['Full Name'])}")

    # --- RIVALS ---
    pre_team_a, pre_team_b, rivalry_notes = [], [], []

    def extract_player(name, pd_pool, pf_pool):
        key = clean_name_key(name)
        pd_pool['MatchKey'] = pd_pool['Full Name'].apply(clean_name_key)
        matches_d = pd_pool[pd_pool['MatchKey'] == key]
        if not matches_d.empty:
            row = matches_d.iloc[0].copy(); row['Position'] = 'D'
            return row, pd_pool.drop(matches_d.index), pf_pool
        
        pf_pool['MatchKey'] = pf_pool['Full Name'].apply(clean_name_key)
        matches_f = pf_pool[pf_pool['MatchKey'] == key]
        if not matches_f.empty:
            row = matches_f.iloc[0].copy(); row['Position'] = 'F'
            return row, pd_pool, pf_pool.drop(matches_f.index)
        return None, pd_pool, pf_pool

    rival_pairs = [("Mike Tonietto", "Jamie Devin"), ("Mark Hicks", "Gary Fera")]
    pair_index = 0 
    for p1, p2 in rival_pairs:
        obj1, pool_d, pool_f = extract_player(p1, pool_d, pool_f)
        obj2, pool_d, pool_f = extract_player(p2, pool_d, pool_f)

        if obj1 is not None and obj2 is not None:
            objs = sorted([obj1, obj2], key=lambda x: x['Score'], reverse=True)
            if pair_index % 2 == 0:
                pre_team_a.append(objs[0]); pre_team_b.append(objs[1])
                rivalry_notes.append(f"Separated {p1} & {p2}")
            else:
                pre_team_b.append(objs[0]); pre_team_a.append(objs[1])
                rivalry_notes.append(f"Separated {p1} & {p2}")
            pair_index += 1
        else:
            if obj1 is not None:
                if obj1['Position'] == 'D': pool_d = pd.concat([pool_d, obj1.to_frame().T])
                else: pool_f = pd.concat([pool_f, obj1.to_frame().T])
            if obj2 is not None:
                if obj2['Position'] == 'D': pool_d = pd.concat([pool_d, obj2.to_frame().T])
                else: pool_f = pd.concat([pool_f, p2_obj.to_frame().T])

    # --- DRAFT ---
    if 'MatchKey' in pool_d.columns: del pool_d['MatchKey']
    if 'MatchKey' in pool_f.columns: del pool_f['MatchKey']
    
    pool_d = pool_d.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])
    pool_f = pool_f.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])

    pre_d = len([p for p in pre_team_a + pre_team_b if p['Position'] == 'D'])
    pre_f = len([p for p in pre_team_a + pre_team_b if p['Position'] == 'F'])
    
    sel_d = pool_d.head(max(0, target_d - pre_d)).copy()
    sel_f = pool_f.head(max(0, target_f - pre_f)).copy()
    cuts_d = pool_d.iloc[len(sel_d):].copy()
    cuts_f = pool_f.iloc[len(sel_f):].copy()

    sel_d['Position'] = 'D'; sel_f['Position'] = 'F'
    da, db = snake_draft(sel_d)
    fa, fb = snake_draft(sel_f)

    # --- COMBINE ---
    final_cols = list(df.columns)
    if 'Position' not in final_cols: final_cols.append('Position')
    if bday_col_name not in final_cols: final_cols.append(bday_col_name)

    def to_df(lst): return pd.DataFrame(lst, columns=final_cols) if lst else pd.DataFrame(columns=final_cols)

    ta = pd.concat([to_df(pre_team_a), da, fa], ignore_index=True)
    tb = pd.concat([to_df(pre_team_b), db, fb], ignore_index=True)
    
    ta = ta.sort_values(by=['Position', 'Full Name']).reset_index(drop=True)
    tb = tb.sort_values(by=['Position', 'Full Name']).reset_index(drop=True)

    # --- DISPLAY ---
    if st.button("Shuffle Teams Again"): st.rerun()
    if rivalry_notes: 
        st.divider()
        for note in rivalry_notes: st.success(f"⚖️ {note}")

    c1, c2 = st.columns(2)
    disp_cols = ['Full Name', 'Position']
    
    with c1:
        st.header("🔴 Red Team")
        st.write(f"Players: {len(ta)} ({len(ta[ta['Position']=='D'])} D / {len(ta[ta['Position']=='F'])} F)")
        st.write(f"Total Score: {ta['Score'].sum():.1f}")
        st.dataframe(ta[disp_cols], hide_index=True)
        
    with c2:
        st.header("⚪ White Team")
        st.write(f"Players: {len(tb)} ({len(tb[tb['Position']=='D'])} D / {len(tb[tb['Position']=='F'])} F)")
        st.write(f"Total Score: {tb['Score'].sum():.1f}")
        st.dataframe(tb[disp_cols], hide_index=True)

    if not cuts_d.empty or not cuts_f.empty:
        st.divider()
        st.error(f"🚫 Undrafted: {len(cuts_d)} D, {len(cuts_f)} F")
        if not cuts_d.empty: st.write(f"D Cuts: {', '.join(cuts_d['Full Name'])}")
        if not cuts_f.empty: st.write(f"F Cuts: {', '.join(cuts_f['Full Name'])}")

    # --- EMAIL ---
    st.divider()
    all_p = pd.concat([ta, tb])
    if not all_p.empty:
        recipients = [e for e in all_p['Email'].unique() if e and str(e).strip()]
        bcc = ",".join(recipients)
        
        # NOTE: bday_msg is already calculated from MASTER list above
        
        body = f"""{bday_msg}Hello everyone,\n\nHere are the rosters for the upcoming game:\n\n{format_team_list(ta, "RED TEAM")}\n{format_team_list(tb, "WHITE TEAM")}\nKeep your sticks on the ice!"""
        st.text_area("Email Draft:", value=body, height=300)

        subj = urllib.parse.quote(f"Bendo Hockey Lineups - {selected_sheet}")
        safe_body = urllib.parse.quote(body)
        safe_bcc = urllib.parse.quote(bcc)
        
        b1, b2, b3 = st.columns(3)
        with b1: st.link_button("📱 Default App", f"mailto:?bcc={safe_bcc}&subject={subj}&body={safe_body}")
        with b2: st.link_button("💻 Gmail Web", f"https://mail.google.com/mail/?view=cm&fs=1&bcc={safe_bcc}&su={subj}&body={safe_body}")
        with b3: st.link_button("🍎 iOS Gmail", f"googlegmail:///co?bcc={safe_bcc}&subject={subj}&body={safe_body}")