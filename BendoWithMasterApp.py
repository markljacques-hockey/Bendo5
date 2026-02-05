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
        # sheet_name=None reads ALL sheets
        # header=1 means Row 2 is the header
        return pd.read_excel(uploaded_file, sheet_name=None, header=1), None
    except Exception as e:
        return None, str(e)

# --- 3. HELPER FUNCTIONS ---
def snake_draft(players):
    players = players.reset_index(drop=True)
    team_a_list, team_b_list = [], []
    for i in range(len(players)):
        player = players.iloc[i]
        if i % 4 == 0 or i % 4 == 3:
            team_a_list.append(player)
        else:
            team_b_list.append(player)
    
    cols = players.columns
    df_a = pd.DataFrame(team_a_list, columns=cols) if team_a_list else pd.DataFrame(columns=cols)
    df_b = pd.DataFrame(team_b_list, columns=cols) if team_b_list else pd.DataFrame(columns=cols)
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
    if pd.isna(name): return ""
    return str(name).lower().replace(" ", "").strip()

def find_col_case_insensitive(df, target_names):
    if isinstance(target_names, str): target_names = [target_names]
    target_names = [t.lower() for t in target_names]
    for col in df.columns:
        if str(col).strip().lower() in target_names: return col
    return None

def get_birthday_message(players_df, bday_col):
    if not bday_col or bday_col not in players_df.columns: return "", []
    
    today = datetime.now()
    
    # SUNDAY LOOKAHEAD LOGIC
    if today.weekday() == 6:
        start_of_week = today + timedelta(days=1)
    else:
        start_of_week = today - timedelta(days=today.weekday())
        
    start_of_week = start_of_week.replace(hour=0, minute=0, second=0, microsecond=0)
    end_of_week = start_of_week + timedelta(days=6)
    end_of_week = end_of_week.replace(hour=23, minute=59, second=59, microsecond=999999)

    celebrants = []
    
    # Work on a copy
    work_df = players_df.copy()
    work_df[bday_col] = pd.to_datetime(work_df[bday_col], errors='coerce')

    for _, row in work_df.iterrows():
        bday = row[bday_col]
        if pd.isna(bday): continue
        try:
            candidates = []
            years = [today.year, today.year - 1, today.year + 1]
            for y in years:
                try: candidates.append(bday.replace(year=y))
                except ValueError: candidates.append(bday.replace(year=y, month=3, day=1))
            
            if any(start_of_week <= c <= end_of_week for c in candidates):
                if 'Full Name' in row: name = row['Full Name']
                elif 'Name' in row: name = row['Name']
                else: name = "Unknown Player"
                
                if name not in celebrants: celebrants.append(name)
        except: continue
            
    if not celebrants: return "", []
    names_str = " and ".join([", ".join(celebrants[:-1]), celebrants[-1]] if len(celebrants) > 2 else celebrants)
    verb = "is" if len(celebrants) == 1 else "are"
    noun = "birthday" if len(celebrants) == 1 else "birthdays"
    return f"🎉 Congratulations to {names_str} who {verb} celebrating their {noun} this week!\n\n", celebrants

# --- 4. MAIN APP ---
st.title("🏒 Hockey Team Generator")
st.write("Upload your Excel player sheet.")

uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls', 'xlsm'])

if uploaded_file is not None:
    # --- LOAD WITH ERROR CHECKING ---
    all_sheets, error_msg = load_excel_data(uploaded_file)
    
    if error_msg:
        st.error(f"❌ Error reading file: {error_msg}")
        st.warning("Troubleshooting:\n1. Ensure file is not open in Excel.\n2. Ensure file is not password protected.")
        st.stop()
        
    all_sheet_names = list(all_sheets.keys())
    master_sheet_name = next((s for s in all_sheet_names if "master" in s.lower()), None)
    
    # FILTER: Strictly remove Master/Instructions/Reference
    block_list = ["master", "instructions", "instruction", "reference", "setup"]
    valid_sheets = [s for s in all_sheet_names if not any(b in s.lower() for b in block_list)]
    
    if not valid_sheets:
        st.error("No valid game sheets found (Hidden: Master, Instructions).")
        st.stop()
        
    selected_sheet = st.selectbox("Select Roster:", valid_sheets)
    df = all_sheets[selected_sheet].copy()

    # --- NORMALIZE COLUMNS ---
    col_name = find_col_case_insensitive(df, ['name', 'full name', 'fullname'])
    col_first = find_col_case_insensitive(df, ['first_name', 'first name'])
    col_last = find_col_case_insensitive(df, ['last_name', 'last name'])
    col_avail = find_col_case_insensitive(df, ['availability', 'avail'])
    col_choice = find_col_case_insensitive(df, ['1st choice', '1stchoice', 'position'])
    col_score = find_col_case_insensitive(df, ['score', 'rating', 'skill'])
    
    if not (col_name or (col_first and col_last)):
        st.error(f"Missing Name columns in '{selected_sheet}'.")
        st.stop()
    if not col_avail or not col_score or not col_choice:
        st.error(f"Missing columns (Availability, Score, or 1st Choice) in '{selected_sheet}'.")
        st.stop()

    df['Availability'] = df[col_avail].astype(str).str.strip().str.upper()
    df['1st Choice'] = df[col_choice].astype(str).str.strip().str.upper()
    df['Score'] = pd.to_numeric(df[col_score], errors='coerce').fillna(0)
    
    if col_name: df['Full Name'] = df[col_name].astype(str).str.strip()
    else: df['Full Name'] = df[col_first].astype(str).str.strip() + ' ' + df[col_last].astype(str).str.strip()
    
    col_2nd = find_col_case_insensitive(df, ['2nd choice', '2ndchoice'])
    col_email = find_col_case_insensitive(df, ['email', 'e-mail'])
    col_reg = find_col_case_insensitive(df, ['reg/spare', 'status'])
    
    df['2nd Choice'] = df[col_2nd].fillna('').astype(str).str.strip().str.upper() if col_2nd else ''
    df['Email'] = df[col_email].fillna('').astype(str).str.strip() if col_email else ''
    df['Reg/Spare'] = df[col_reg] if col_reg else 'R'

    # --- MASTER LOOKUP (BIRTHDAYS) ---
    bday_col_name = 'B-day'
    birthday_source = df # Default to daily sheet
    mapped_count = 0
    
    if master_sheet_name:
        try:
            df_master = all_sheets[master_sheet_name].copy()
            m_name = find_col_case_insensitive(df_master, ['name', 'full name'])
            m_first = find_col_case_insensitive(df_master, ['first_name', 'first name'])
            m_last = find_col_case_insensitive(df_master, ['last_name', 'last name'])
            m_bday = find_col_case_insensitive(df_master, ['b-day', 'bday', 'birthday', 'dob'])
            
            if (m_name or (m_first and m_last)) and m_bday:
                # Prepare Master Key
                if m_name: df_master['Key'] = df_master[m_name].apply(clean_name_key)
                else: df_master['Key'] = (df_master[m_first].astype(str) + df_master[m_last].astype(str)).apply(clean_name_key)
                
                # Construct Full Name for Greeting
                if m_name: df_master['Full Name'] = df_master[m_name]
                else: df_master['Full Name'] = df_master[m_first].astype(str) + ' ' + df_master[m_last].astype(str)

                df_master[bday_col_name] = pd.to_datetime(df_master[m_bday], errors='coerce')
                
                # USE MASTER FOR GREETING SOURCE
                birthday_source = df_master 
                
                # Map to Daily for Debugging/Reference
                bday_map = df_master.set_index('Key')[bday_col_name].to_dict()
                df['Key'] = df['Full Name'].apply(clean_name_key)
                df[bday_col_name] = df['Key'].map(bday_map)
                mapped_count = df[bday_col_name].notna().sum()
            else:
                st.warning("Master sheet found but missing Name/B-day columns.")
        except Exception as e:
            st.warning(f"Master sheet error: {e}")
    else:
        # Fallback to daily
        daily_bday = find_col_case_insensitive(df, ['b-day', 'bday', 'birthday'])
        if daily_bday: 
            bday_col_name = daily_bday
            df[bday_col_name] = pd.to_datetime(df[bday_col_name], errors='coerce')
            mapped_count = df[bday_col_name].notna().sum()

    # --- FILTER ---
    available = df[df['Availability'].str.startswith('Y')].copy()
    if available.empty:
        st.error("No players marked 'Yes'.")
        st.stop()
    
    # --- TARGETS ---
    total = len(available)
    BASE_F, BASE_D, MIN_D_CRITICAL = 12, 8, 6

    if total <= 20: target_f, target_d = BASE_F, BASE_D
    else:
        ex = total - 20
        add_f = min(ex, 6); ex -= add_f
        add_d = min(ex, 4)
        target_f, target_d = BASE_F + add_f, BASE_D + add_d
    
    st.info(f"**Strategy ({selected_sheet}):** {total} players -> {target_f} F / {target_d} D")
    
    # --- BIRTHDAY CHECK (Uses birthday_source) ---
    bday_msg, bday_names = get_birthday_message(birthday_source, bday_col_name)
    with st.expander("🎂 Birthday Debugger", expanded=True):
        if bday_names: st.success(f"**Birthdays Found:** {', '.join(bday_names)}")
        else: st.info("No birthdays found this week (checked Master list).")
        st.caption(f"Checked {len(birthday_source)} players in source list.")

    # --- LOGIC ---
    available = available.sample(frac=1).reset_index(drop=True)
    available['Status_Rank'] = available['Reg/Spare'].apply(lambda x: 0 if str(x).strip().upper() == 'R' else 1)
    available = available.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])

    pool_d = available[available['1st Choice'] == 'D'].copy()
    pool_f = available[available['1st Choice'] == 'F'].copy()

    # --- BALANCER LOGIC (Surplus Checks & Named Alerts) ---
    # 1. Critical D Shortage (Must have at least 6 D)
    if len(pool_d) < MIN_D_CRITICAL:
        needed = MIN_D_CRITICAL - len(pool_d)
        candidates = pool_f[pool_f['2nd Choice'] == 'D']
        if not candidates.empty:
            conv = candidates.head(needed)
            pool_d = pd.concat([pool_d, conv])
            pool_f = pool_f.drop(conv.index)
            st.warning(f"⚠️ Critical D Shortage (<6): Moved {len(conv)} F -> D: **{', '.join(conv['Full Name'])}**")

    # 2. Forward Shortage (Only take from D if D has surplus > 6)
    if len(pool_f) < target_f:
        needed = target_f - len(pool_f)
        surplus_d = len(pool_d) - MIN_D_CRITICAL # Keep D above 6
        
        if surplus_d > 0:
            can_take = min(needed, surplus_d)
            candidates = pool_d[pool_d['2nd Choice'] == 'F']
            if not candidates.empty:
                conv = candidates.head(can_take)
                pool_f = pd.concat([pool_f, conv])
                pool_d = pool_d.drop(conv.index)
                st.info(f"Balanced Forwards: Moved {len(conv)} D -> F: **{', '.join(conv['Full Name'])}**")

    # 3. Defense Shortage (Only take from F if F has surplus > Target)
    if len(pool_d) < target_d:
        d_shortage = target_d - len(pool_d)
        surplus_f = len(pool_f) - target_f # Only take if F has more than they need
        
        if surplus_f > 0:
            amount_to_move = min(d_shortage, surplus_f)
            candidates = pool_f[pool_f['2nd Choice'] == 'D']
            if not candidates.empty:
                conv = candidates.head(amount_to_move)
                pool_d = pd.concat([pool_d, conv])
                pool_f = pool_f.drop(conv.index)
                st.info(f"Balanced Defense: Moved {len(conv)} F -> D: **{', '.join(conv['Full Name'])}**")

    # --- RIVALS ---
    pre_a, pre_b, rival_logs = [], [], []
    def extract(nm, pd_pool, pf_pool):
        key = clean_name_key(nm)
        pd_pool['K'] = pd_pool['Full Name'].apply(clean_name_key)
        pf_pool['K'] = pf_pool['Full Name'].apply(clean_name_key)
        md, mf = pd_pool[pd_pool['K']==key], pf_pool[pf_pool['K']==key]
        if not md.empty: return md.iloc[0].copy(), pd_pool.drop(md.index), pf_pool
        if not mf.empty: return mf.iloc[0].copy(), pd_pool, pf_pool.drop(mf.index)
        return None, pd_pool, pf_pool

    rivals = [("Mike Tonietto", "Jamie Devin"), ("Mark Hicks", "Gary Fera")]
    pidx = 0
    for p1, p2 in rivals:
        o1, pool_d, pool_f = extract(p1, pool_d, pool_f)
        o2, pool_d, pool_f = extract(p2, pool_d, pool_f)
        if o1 is not None and o2 is not None:
            srt = sorted([o1, o2], key=lambda x: x['Score'], reverse=True)
            if pidx % 2 == 0: pre_a.append(srt[0]); pre_b.append(srt[1])
            else: pre_b.append(srt[0]); pre_a.append(srt[1])
            rival_logs.append(f"Separated {p1} & {p2}")
            pidx += 1
        else: # Return if only 1 present
            if o1 is not None: 
                if o1['1st Choice']=='D': pool_d = pd.concat([pool_d, o1.to_frame().T])
                else: pool_f = pd.concat([pool_f, o1.to_frame().T])
            if o2 is not None:
                if o2['1st Choice']=='D': pool_d = pd.concat([pool_d, o2.to_frame().T])
                else: pool_f = pd.concat([pool_f, o2.to_frame().T])

    # Draft
    if 'K' in pool_d.columns: del pool_d['K']
    if 'K' in pool_f.columns: del pool_f['K']
    
    pool_d = pool_d.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])
    pool_f = pool_f.sort_values(by=['Status_Rank', 'Score'], ascending=[True, False])

    pd_cnt = len([p for p in pre_a+pre_b if p['1st Choice']=='D'])
    pf_cnt = len([p for p in pre_a+pre_b if p['1st Choice']=='F'])
    
    sel_d = pool_d.head(max(0, target_d - pd_cnt)).copy()
    sel_f = pool_f.head(max(0, target_f - pf_cnt)).copy()
    cuts_d = pool_d.iloc[len(sel_d):].copy()
    cuts_f = pool_f.iloc[len(sel_f):].copy()

    sel_d['Position'] = 'D'; sel_f['Position'] = 'F'
    for p in pre_a+pre_b: p['Position'] = p['1st Choice']

    da, db = snake_draft(sel_d)
    fa, fb = snake_draft(sel_f)

    # Combine
    cols = ['Full Name', 'Position', 'Score', 'Email']
    def to_df(l): return pd.DataFrame(l) if l else pd.DataFrame(columns=cols)
    
    ta = pd.concat([to_df(pre_a), da, fa], ignore_index=True)
    tb = pd.concat([to_df(pre_b), db, fb], ignore_index=True)
    
    # Sort
    ta = ta.sort_values(by=['Position', 'Full Name']).reset_index(drop=True)
    tb = tb.sort_values(by=['Position', 'Full Name']).reset_index(drop=True)

    # UI
    if st.button("Shuffle"): st.rerun()
    if rival_logs: 
        st.divider()
        for r in rival_logs: st.success(f"⚖️ {r}")

    # --- SCORE CALCULATIONS ---
    common_count = min(len(ta), len(tb))
    total_a = ta['Score'].sum()
    total_b = tb['Score'].sum()
    fair_a = get_top_n_score(ta, common_count)
    fair_b = get_top_n_score(tb, common_count)

    c1, c2 = st.columns(2)
    with c1:
        st.header("🔴 Red Team")
        st.write(f"**Total Score:** {total_a:.1f}")
        if len(ta) != len(tb): st.write(f"**Top {common_count} Score:** {fair_a:.1f}")
        st.write(f"Players: {len(ta)} ({len(ta[ta['Position']=='D'])} D / {len(ta[ta['Position']=='F'])} F)")
        st.dataframe(ta[['Full Name', 'Position']], hide_index=True)
    with c2:
        st.header("⚪ White Team")
        st.write(f"**Total Score:** {total_b:.1f}")
        if len(ta) != len(tb): st.write(f"**Top {common_count} Score:** {fair_b:.1f}")
        st.write(f"Players: {len(tb)} ({len(tb[tb['Position']=='D'])} D / {len(tb[tb['Position']=='F'])} F)")
        st.dataframe(tb[['Full Name', 'Position']], hide_index=True)

    if not cuts_d.empty or not cuts_f.empty:
        st.error(f"🚫 Cuts: {len(cuts_d)} D, {len(cuts_f)} F")
        if not cuts_d.empty: st.write(f"D: {', '.join(cuts_d['Full Name'])}")
        if not cuts_f.empty: st.write(f"F: {', '.join(cuts_f['Full Name'])}")

    st.divider()
    all_p = pd.concat([ta, tb])
    if not all_p.empty:
        # Email
        recipients = [e for e in all_p['Email'].unique() if pd.notna(e) and str(e).strip()]
        bcc = ",".join(recipients)
        body = f"""{bday_msg}Hello everyone,\n\nHere are the rosters for the upcoming game:\n\n{format_team_list(ta, "RED TEAM")}\n{format_team_list(tb, "WHITE TEAM")}\nKeep your sticks on the ice!"""
        
        st.text_area("Email Draft", value=body, height=300)
        subj = urllib.parse.quote(f"Bendo Hockey Lineups - {selected_sheet}")
        safe_body = urllib.parse.quote(body)
        safe_bcc = urllib.parse.quote(bcc)
        
        c1, c2, c3 = st.columns(3)
        c1.link_button("📱 Default App", f"mailto:?bcc={safe_bcc}&subject={subj}&body={safe_body}")
        c2.link_button("💻 Gmail Web", f"https://mail.google.com/mail/?view=cm&fs=1&bcc={safe_bcc}&su={subj}&body={safe_body}")
        c3.link_button("🍎 iOS Gmail", f"googlegmail:///co?bcc={safe_bcc}&subject={subj}&body={safe_body}")