import streamlit as st
import pandas as pd
from datetime import datetime
import io
import requests
import json

# --- Configuration & Constants ---
PAGE_TITLE = "Kingspan Sales Report Generator"
PAGE_ICON = "📊"

AM_MAPPING = {
    'nicolaas': 'Nicolaas',
    'pieter': 'Pieter',
    'arne': 'Arne',
    'simon': 'Simon',
    'omayra': 'Omayra',
    'ayoub': 'Ayoub'
}

VALID_AMS = {'Pieter', 'Arne', 'Nicolaas', 'Simon'}

# Translations
TRANSLATIONS = {
    'NL': {
        'title': "[Sales + Ops] Rapport Kingspanstore.be (BA webshop)",
        'intro_greeting': "Dag collega's,",
        'intro_text': "Hierbij de update van de BE-NL Shopify-store resultaten voor **{period}**. We vergelijken deze periode met de voorgaande en kijken naar de totale groei sinds de start.",
        'total_since_start': "Totaal sinds start (15 april 2025)",
        'total_revenue': "Totale Omzet",
        'total_orders': "Aantal Orders",
        'avg_order_value': "Gem. Orderwaarde",
        'period_results': "Resultaten {period}",
        'revenue_period': "Omzet ({period})",
        'orders_period': "Orders ({period})",
        'key_insights': "Belangrijkste inzichten",
        'new_customers': "Nieuwe Klanten",
        'new_customers_text': "In {period} mochten we **{count}** nieuwe klanten verwelkomen.",
        'customer_db': "Klantendatabase",
        'customer_db_text': "{active} actieve klanten (die al besteld hebben) vs {inactive} inactieve klanten.",
        'am_stats': "Accountmanagers",
        'am_stats_text': "{with_am} klanten hebben een AM gekoppeld, {without_am} nog niet.",
        'challenges_opps': "Uitdagingen & Kansen",
        'monthly_evolution': "Maandelijkse evolutie",
        'month': "Maand",
        'revenue': "Omzet",
        'orders': "Orders",
        'am_revenue': "Omzet per accountmanager ({period})",
        'top_returning': "Top 5 terugkerende klanten ({period})",
        'customer': "Klant",
        'top_revenue': "Top 5 klanten op basis van omzet ({period})",
        'top_products': "Top 5 meest verkochte producten ({period})",
        'product': "Product",
        'quantity': "Aantal",
        'customer_insights': "Klanteninzichten",
        'category': "Categorie",
        'percentage': "Percentage",
        'total_customers': "Totaal aantal klanten",
        'active_customers': "Actieve klanten (min. 1 order)",
        'inactive_customers': "Inactieve klanten (0 orders)",
        'customers_with_am': "Klanten met Accountmanager",
        'customers_without_am': "Klanten zonder Accountmanager",
        'new_registrations': "Nieuwe registraties ({period})",
        'footer_contact': "Bij vragen of aanvullende inzichten: stuur maar een mail of bericht op Teams",
        'footer_rights': "Kingspan Light + Air België | Ayoub en Omayra"
    },
    'FR': {
        'title': "[Sales + Ops] Rapport Kingspanstore.be (BA webshop)",
        'intro_greeting': "Bonjour collègues,",
        'intro_text': "Voici la mise à jour des résultats du Shopify-store BE-NL pour **{period}**. Nous comparons cette période avec la précédente et examinons la croissance totale depuis le début.",
        'total_since_start': "Total depuis le début (15 avril 2025)",
        'total_revenue': "Chiffre d'affaires total",
        'total_orders': "Nombre de commandes",
        'avg_order_value': "Panier moyen",
        'period_results': "Résultats {period}",
        'revenue_period': "CA ({period})",
        'orders_period': "Commandes ({period})",
        'key_insights': "Points clés",
        'new_customers': "Nouveaux Clients",
        'new_customers_text': "En {period}, nous avons accueilli **{count}** nouveaux clients.",
        'customer_db': "Base de données clients",
        'customer_db_text': "{active} clients actifs (ayant commandé) vs {inactive} clients inactifs.",
        'am_stats': "Account Managers",
        'am_stats_text': "{with_am} clients ont un AM lié, {without_am} pas encore.",
        'challenges_opps': "Défis & Opportunités",
        'monthly_evolution': "Évolution mensuelle",
        'month': "Mois",
        'revenue': "CA",
        'orders': "Commandes",
        'am_revenue': "CA par Account Manager ({period})",
        'top_returning': "Top 5 clients récurrents ({period})",
        'customer': "Client",
        'top_revenue': "Top 5 clients par CA ({period})",
        'top_products': "Top 5 produits vendus ({period})",
        'product': "Produit",
        'quantity': "Quantité",
        'customer_insights': "Aperçu Clients",
        'category': "Catégorie",
        'percentage': "Pourcentage",
        'total_customers': "Nombre total de clients",
        'active_customers': "Clients actifs (min. 1 commande)",
        'inactive_customers': "Clients inactifs (0 commande)",
        'customers_with_am': "Clients avec Account Manager",
        'customers_without_am': "Clients sans Account Manager",
        'new_registrations': "Nouvelles inscriptions ({period})",
        'footer_contact': "Pour toute question ou information complémentaire : envoyez un mail ou un message sur Teams",
        'footer_rights': "Kingspan Light + Air Belgique | Ayoub et Omayra"
    },
    'EN': {
        'title': "[Sales + Ops] Report Kingspanstore.be (BA webshop)",
        'intro_greeting': "Hi colleagues,",
        'intro_text': "Here is the update of the BE-NL Shopify-store results for **{period}**. We compare this period with the previous one and look at the total growth since the start.",
        'total_since_start': "Total since start (April 15, 2025)",
        'total_revenue': "Total Revenue",
        'total_orders': "Number of Orders",
        'avg_order_value': "Avg. Order Value",
        'period_results': "Results {period}",
        'revenue_period': "Revenue ({period})",
        'orders_period': "Orders ({period})",
        'key_insights': "Key Insights",
        'new_customers': "New Customers",
        'new_customers_text': "In {period}, we welcomed **{count}** new customers.",
        'customer_db': "Customer Database",
        'customer_db_text': "{active} active customers (who ordered) vs {inactive} inactive customers.",
        'am_stats': "Account Managers",
        'am_stats_text': "{with_am} customers have an AM linked, {without_am} do not.",
        'challenges_opps': "Challenges & Opportunities",
        'monthly_evolution': "Monthly Evolution",
        'month': "Month",
        'revenue': "Revenue",
        'orders': "Orders",
        'am_revenue': "Revenue per Account Manager ({period})",
        'top_returning': "Top 5 Returning Customers ({period})",
        'customer': "Customer",
        'top_revenue': "Top 5 Customers by Revenue ({period})",
        'top_products': "Top 5 Best Selling Products ({period})",
        'product': "Product",
        'quantity': "Quantity",
        'customer_insights': "Customer Insights",
        'category': "Category",
        'percentage': "Percentage",
        'total_customers': "Total Customers",
        'active_customers': "Active Customers (min. 1 order)",
        'inactive_customers': "Inactive Customers (0 orders)",
        'customers_with_am': "Customers with Account Manager",
        'customers_without_am': "Customers without Account Manager",
        'new_registrations': "New Registrations ({period})",
        'footer_contact': "For questions or additional insights: send an email or message on Teams",
        'footer_rights': "Kingspan Light + Air Belgium | Ayoub and Omayra"
    }
}

MONTH_NAMES = {
    'NL': {1: 'januari', 2: 'februari', 3: 'maart', 4: 'april', 5: 'mei', 6: 'juni', 7: 'juli', 8: 'augustus', 9: 'september', 10: 'oktober', 11: 'november', 12: 'december'},
    'FR': {1: 'janvier', 2: 'février', 3: 'mars', 4: 'avril', 5: 'mai', 6: 'juin', 7: 'juillet', 8: 'août', 9: 'septembre', 10: 'octobre', 11: 'novembre', 12: 'décembre'},
    'EN': {1: 'January', 2: 'February', 3: 'March', 4: 'April', 5: 'May', 6: 'June', 7: 'July', 8: 'August', 9: 'September', 10: 'October', 11: 'November', 12: 'December'}
}

# --- Helper Functions ---

def fmt_eur(val):
    return f"€ {val:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

def extract_am_from_tags(tags_str):
    if not isinstance(tags_str, str): return None
    tags = [t.strip().lower() for t in tags_str.split(',')]
    for tag in tags:
        if tag.startswith('am:'):
            email_part = tag.split(':', 1)[1]
            for key, name in AM_MAPPING.items():
                if key in email_part:
                    return name
    return None

def is_test_account(email):
    if not isinstance(email, str): return False
    email = email.lower()
    if 'ayoubmohyi' in email or 'ayoub.mohyi' in email: return True
    if 'inscrlab.com' in email: return True
    if 'protonmail' in email: return True
    if 'kingspan.com' in email: return True # Added based on previous context
    return False

# --- Main App ---

st.set_page_config(page_title=PAGE_TITLE, page_icon=PAGE_ICON, layout="wide")

st.title(f"{PAGE_ICON} {PAGE_TITLE}")
st.markdown("Upload your Shopify exports to generate the HTML sales report.")

# --- Sidebar: Settings ---
st.sidebar.header("1. Settings")
language = st.sidebar.selectbox("Language / Taal / Langue", ["NL", "FR", "EN"], index=0)
t = TRANSLATIONS[language]

st.sidebar.header("2. Data Upload")
orders_file = st.sidebar.file_uploader("Export All Orders (Excel)", type=['xlsx'])
customers_file = st.sidebar.file_uploader("Export Customers (Excel)", type=['xlsx'])

st.sidebar.header("3. Period Selection")
current_year = datetime.now().year
selected_year = st.sidebar.number_input("Year", min_value=2024, max_value=2030, value=current_year)
selected_months = st.sidebar.multiselect("Months", range(1, 13), default=[datetime.now().month - 1] if datetime.now().month > 1 else [12], format_func=lambda x: MONTH_NAMES[language][x])

st.sidebar.header("4. AI Insights (Optional)")
make_webhook_url = st.sidebar.text_input("Make.com Webhook URL", placeholder="https://hook.eu1.make.com/...")
generate_ai_btn = st.sidebar.button("✨ Generate Insights with AI")

# --- Main Logic ---

if orders_file and customers_file and selected_months:
    try:
        # Load Data
        with st.spinner('Processing data...'):
            df_orders = pd.read_excel(orders_file)
            df_customers = pd.read_excel(customers_file)
            
            # --- Preprocessing Orders ---
            # Standardize columns (using the ones identified previously)
            # Note: Pandas reads headers. We assume standard Shopify export headers.
            # We need to map the column names or find them dynamically if they change.
            # For robustness, let's try to find columns by keyword if exact match fails.
            
            def get_col(df, keywords):
                for col in df.columns:
                    if any(k.lower() in col.lower() for k in keywords):
                        return col
                return None

            col_id = get_col(df_orders, ['Name', 'Order ID']) # Usually 'Name' is #1001
            col_email = get_col(df_orders, ['Email'])
            col_created = get_col(df_orders, ['Created at'])
            col_total = get_col(df_orders, ['Total'])
            col_cust_tags = get_col(df_orders, ['Customer Tags']) # Important for AM
            col_line_type = get_col(df_orders, ['Lineitem type']) # To distinguish products
            col_prod_title = get_col(df_orders, ['Lineitem name', 'Product'])
            col_qty = get_col(df_orders, ['Lineitem quantity'])
            col_cancelled = get_col(df_orders, ['Cancelled at'])

            # Convert dates
            # Force UTC to ensure we get a valid datetime series (not object), then strip timezone
            df_orders['date'] = pd.to_datetime(df_orders[col_created], errors='coerce', utc=True)
            df_orders['date'] = df_orders['date'].dt.tz_localize(None)
            
            # Filter Valid Orders (Start Date)
            start_date = pd.Timestamp(2025, 4, 15)
            valid_orders = df_orders[
                (df_orders['date'] >= start_date) & 
                (df_orders[col_cancelled].isna())
            ].copy()
            
            # Extract AM
            valid_orders['am'] = valid_orders[col_cust_tags].apply(extract_am_from_tags)
            
            # Filter Period Orders
            period_orders = valid_orders[
                (valid_orders['date'].dt.year == selected_year) & 
                (valid_orders['date'].dt.month.isin(selected_months))
            ].copy()
            
            # --- Preprocessing Customers ---
            col_c_email = get_col(df_customers, ['Email'])
            col_c_created = get_col(df_customers, ['Created at'])
            col_c_orders = get_col(df_customers, ['Total Orders'])
            col_c_tags = get_col(df_customers, ['Tags'])
            
            df_customers['created_at'] = pd.to_datetime(df_customers[col_c_created])
            
            # Deduplicate customers (keep latest or aggregate) - Shopify export usually has one row per customer
            # But if multiple addresses, might have duplicates. Group by Email.
            
            # Filter Test Accounts
            df_customers['is_test'] = df_customers[col_c_email].apply(is_test_account)
            real_customers = df_customers[~df_customers['is_test']].copy()
            
            # Calculate Customer Stats
            total_customers_count = len(real_customers)
            active_customers_count = len(real_customers[real_customers[col_c_orders] > 0])
            inactive_customers_count = total_customers_count - active_customers_count
            
            real_customers['am'] = real_customers[col_c_tags].apply(extract_am_from_tags)
            customers_with_am_count = len(real_customers[real_customers['am'].isin(VALID_AMS)])
            customers_without_am_count = total_customers_count - customers_with_am_count
            
            # New Customers in Period
            new_customers_period = real_customers[
                (real_customers['created_at'].dt.year == selected_year) & 
                (real_customers['created_at'].dt.month.isin(selected_months))
            ]
            new_customers_count = len(new_customers_period)

            # --- KPI Calculations ---
            # Total Since Start
            # Need to deduplicate orders (one row per line item in export?)
            # Usually Shopify 'Export All Orders' has multiple rows per order.
            # We need to group by Order Name/ID to get totals.
            
            unique_orders_all = valid_orders.drop_duplicates(subset=[col_id])
            launch_revenue = unique_orders_all[col_total].sum()
            launch_orders = len(unique_orders_all)
            launch_avg = launch_revenue / launch_orders if launch_orders else 0
            
            unique_orders_period = period_orders.drop_duplicates(subset=[col_id])
            period_revenue = unique_orders_period[col_total].sum()
            period_orders_count = len(unique_orders_period)
            period_avg = period_revenue / period_orders_count if period_orders_count else 0
            
            # Monthly Evolution (Last 3 months relevant to selection)
            # Get all months in data
            monthly_stats = unique_orders_all.groupby([unique_orders_all['date'].dt.year, unique_orders_all['date'].dt.month])[col_total].agg(['count', 'sum']).sort_index()
            
            # AM Performance
            am_stats = unique_orders_period.groupby('am')[col_total].agg(['count', 'sum']).reset_index()
            # Add 'No AM'
            no_am_rev = unique_orders_period[unique_orders_period['am'].isna()][col_total].sum()
            no_am_count = len(unique_orders_period[unique_orders_period['am'].isna()])
            if no_am_count > 0:
                am_stats = pd.concat([am_stats, pd.DataFrame({'am': ['Geen AM'], 'count': [no_am_count], 'sum': [no_am_rev]})])
            am_stats = am_stats.sort_values('sum', ascending=False)
            
            # Top Products
            # Filter for line items only
            # Assuming 'Lineitem type' exists or we infer it.
            # If not, usually rows with Product Name are line items.
            prod_stats = period_orders.groupby(col_prod_title)[col_qty].sum().sort_values(ascending=False).head(5)
            
            # Top Customers
            cust_rev = unique_orders_period.groupby(col_email)[col_total].sum().sort_values(ascending=False).head(5)
            cust_freq = unique_orders_period.groupby(col_email).size().sort_values(ascending=False).head(5)

        # --- Input Fields for Text ---
        st.subheader("5. Report Content")
        
        period_str = ", ".join([MONTH_NAMES[language][m] for m in selected_months]) + f" {selected_year}"
        
        default_challenges = "Onze grootste uitdaging (en kans!) is het verbreden van onze klantenbasis..."
        
        # AI Generation Handler
        if generate_ai_btn and make_webhook_url:
            with st.spinner("Asking AI..."):
                payload = {
                    "period": period_str,
                    "revenue": period_revenue,
                    "orders": period_orders_count,
                    "new_customers": new_customers_count,
                    "top_product": prod_stats.index[0] if not prod_stats.empty else "N/A",
                    "language": language
                }
                try:
                    response = requests.post(make_webhook_url, json=payload)
                    if response.status_code == 200:
                        ai_text = response.text # Assuming Make returns raw text
                        default_challenges = ai_text
                        st.success("AI Insights Generated!")
                    else:
                        st.error("Failed to contact Make.com")
                except Exception as e:
                    st.error(f"Error: {e}")

        challenges_text = st.text_area("Challenges & Opportunities Text", value=default_challenges, height=100)
        
        # --- HTML Generation ---
        if st.button("Generate HTML Report"):
            
            # Construct HTML
            html_content = f"""<!doctype html>
<html lang="{language.lower()}">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>{t['title']}</title>
  <style type="text/css">
    body {{ margin: 0; padding: 0; background-color: #fefefe; font-family: "Segoe UI", Tahoma, sans-serif; }}
    table {{ border-collapse: collapse; border-spacing: 0; }}
    img {{ border: 0; display: block; }}
    .email-wrapper {{ width: 100%; background-color: #fefefe; }}
    .email-container {{ width: 100%; max-width: 720px; margin: 0 auto; background-color: #ffffff; }}
    .header {{ border-bottom: 4px solid #004289; background-color: #ffffff; padding: 25px 0; }}
    .logo-container {{ width: 600px; margin: 0 auto; text-align: center; }}
    .logo {{ width: 140px; height: 85px; }}
    .content {{ background-color: #ffffff; padding: 40px 0; }}
    .content-inner {{ width: 600px; margin: 0 auto; padding: 30px; color: #3c3c3b; }}
    .content h1 {{ margin: 0 0 20px; font-size: 24px; font-weight: 600; color: #004289; }}
    .content h2 {{ margin: 25px 0 15px; font-size: 18px; font-weight: 600; color: #004289; }}
    .content p {{ line-height: 1.5; margin-bottom: 20px; font-size: 16px; }}
    .kpis {{ width: 100%; margin: 20px 0; }}
    .kpi-cell {{ border: 1px solid #e6e8ee; padding: 14px; text-align: center; background-color: #f8f9fb; font-size: 13px; }}
    .kpi-value {{ font-size: 20px; font-weight: 700; color: #004289; margin-bottom: 5px; }}
    .kpi-label {{ font-size: 11px; color: #586375; font-weight: 400; line-height: 1.3; text-transform: uppercase; letter-spacing: 0.5px; }}
    .alert {{ border-left: 4px solid #c09a5d; background-color: #f8f8f8; padding: 20px; margin: 20px 0; border-radius: 4px; }}
    .alert ul {{ margin: 0; padding-left: 18px; }}
    .alert li {{ margin-bottom: 8px; }}
    .data-table {{ width: 100%; margin: 20px 0; border: 1px solid #e6e8ee; }}
    .data-table th {{ background-color: #004289; color: #ffffff; padding: 12px; text-align: left; font-weight: 700; font-size: 13px; text-transform: uppercase; letter-spacing: 0.4px; }}
    .data-table td {{ padding: 12px; border-bottom: 1px solid #e6e8ee; font-size: 14px; }}
    .data-table td.number {{ text-align: right; font-weight: 600; }}
    .footer {{ background-color: #142840; color: #ffffff; padding: 35px 0; }}
    .footer-inner {{ width: 600px; margin: 0 auto; text-align: center; }}
    .footer p {{ margin: 0 0 15px; font-size: 14px; }}
    .footer a {{ color: #c09a5d; text-decoration: none; }}
  </style>
</head>
<body style="margin:0; padding:0; background-color:#fefefe; font-family:&quot;Segoe UI&quot;, Tahoma, sans-serif;">
  <div class="email-wrapper" style="width:100%; background-color:#fefefe;">
    <table class="email-container" role="presentation" cellpadding="0" cellspacing="0" align="center" width="720" style="width:100%; max-width:720px; background-color:#ffffff; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
      <tr>
        <td class="header" style="border-bottom:4px solid #004289; background-color:#ffffff; padding:25px 0;">
          <div class="logo-container" style="width:600px; margin:0 auto; text-align:center;">
            <a href="https://kingspanstore.be" style="text-decoration: none;">
              <img src="https://cdn.shopify.com/s/files/1/0673/3732/2646/files/Kingspan_Light___Air_Logo_PNG_Image_UK_EN_5fe4be0c-74cc-4e07-88fc-d7bc8373edf9.png" alt="Kingspan Light + Air Logo" class="logo" width="140" height="85" style="display:block; width:140px; height:85px; border:0;">
            </a>
          </div>
        </td>
      </tr>
      <tr>
        <td class="content" style="background-color:#ffffff; padding:40px 0;">
          <div class="content-inner" style="width:100%; max-width:600px; margin:0 auto; padding:0;">
            <table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" style="width:100%; max-width:600px; border-collapse:collapse; mso-table-lspace:0pt; mso-table-rspace:0pt;">
              <tr>
                <td style="padding:30px; color:#3c3c3b; font-family:'Segoe UI', Tahoma, sans-serif; font-size:16px; line-height:24px; mso-line-height-rule:exactly;">
            <h1 style="margin:0 0 20px; font-size:24px; font-weight:600; color:#004289;">{t['title']} - {period_str}</h1>
            <p>{t['intro_greeting']}</p>
            <p>{t['intro_text'].format(period=period_str)}</p>
            
            <h2 style="margin:25px 0 15px; font-size:18px; font-weight:600; color:#004289;">&#128202; {t['total_since_start']}</h2>
            <table class="kpis" role="presentation" cellpadding="0" cellspacing="0" style="width:100%; margin:10px 0 20px; border-collapse:separate; border-spacing:0; mso-table-lspace:0pt; mso-table-rspace:0pt;">
              <tr>
                <td class="kpi-cell" style="border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:13px; padding:14px; text-align:center; background-color:#f8f9fb;">
                  <div class="kpi-value" style="font-size:20px; font-weight:700; color:#004289; margin-bottom:5px;">{fmt_eur(launch_revenue)}</div>
                  <div class="kpi-label" style="font-size:11px; color:#586375; font-weight:400; line-height:1.3; text-transform:uppercase; letter-spacing:0.5px;">{t['total_revenue']}</div>
                </td>
                <td class="kpi-cell" style="border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:13px; padding:14px; text-align:center; background-color:#f8f9fb;">
                  <div class="kpi-value" style="font-size:20px; font-weight:700; color:#004289; margin-bottom:5px;">{launch_orders}</div>
                  <div class="kpi-label" style="font-size:11px; color:#586375; font-weight:400; line-height:1.3; text-transform:uppercase; letter-spacing:0.5px;">{t['total_orders']}</div>
                </td>
                <td class="kpi-cell" style="border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:13px; padding:14px; text-align:center; background-color:#f8f9fb;">
                  <div class="kpi-value" style="font-size:20px; font-weight:700; color:#004289; margin-bottom:5px;">{fmt_eur(launch_avg)}</div>
                  <div class="kpi-label" style="font-size:11px; color:#586375; font-weight:400; line-height:1.3; text-transform:uppercase; letter-spacing:0.5px;">{t['avg_order_value']}</div>
                </td>
              </tr>
            </table>

            <h2 style="margin:25px 0 15px; font-size:18px; font-weight:600; color:#004289;">&#128197; {t['period_results'].format(period=period_str)}</h2>
            <table class="kpis" role="presentation" cellpadding="0" cellspacing="0" style="width:100%; margin:10px 0 20px; border-collapse:separate; border-spacing:0; mso-table-lspace:0pt; mso-table-rspace:0pt;">
              <tr>
                <td class="kpi-cell" style="border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:13px; padding:14px; text-align:center; background-color:#f8f9fb;">
                  <div class="kpi-value" style="font-size:20px; font-weight:700; color:#004289; margin-bottom:5px;">{fmt_eur(period_revenue)}</div>
                  <div class="kpi-label" style="font-size:11px; color:#586375; font-weight:400; line-height:1.3; text-transform:uppercase; letter-spacing:0.5px;">{t['revenue_period'].format(period=period_str)}</div>
                </td>
                <td class="kpi-cell" style="border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:13px; padding:14px; text-align:center; background-color:#f8f9fb;">
                  <div class="kpi-value" style="font-size:20px; font-weight:700; color:#004289; margin-bottom:5px;">{period_orders_count}</div>
                  <div class="kpi-label" style="font-size:11px; color:#586375; font-weight:400; line-height:1.3; text-transform:uppercase; letter-spacing:0.5px;">{t['orders_period'].format(period=period_str)}</div>
                </td>
                <td class="kpi-cell" style="border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:13px; padding:14px; text-align:center; background-color:#f8f9fb;">
                  <div class="kpi-value" style="font-size:20px; font-weight:700; color:#004289; margin-bottom:5px;">{fmt_eur(period_avg)}</div>
                  <div class="kpi-label" style="font-size:11px; color:#586375; font-weight:400; line-height:1.3; text-transform:uppercase; letter-spacing:0.5px;">{t['avg_order_value']}</div>
                </td>
              </tr>
            </table>

            <div class="alert" style="border-left:4px solid #c09a5d; background-color:#f8f8f8; padding:20px; margin:20px 0; mso-margin-top-alt:20px; mso-margin-bottom-alt:20px;">
              <h2 style="margin:25px 0 15px; font-size:18px; font-weight:600; color:#004289;">&#128202; {t['key_insights']}</h2>
              <ul style="margin:0; padding-left:18px;">
                <li><strong>{t['new_customers']}:</strong> {t['new_customers_text'].format(period=period_str, count=new_customers_count)}</li>
                <li><strong>{t['customer_db']}:</strong> {t['customer_db_text'].format(active=active_customers_count, inactive=inactive_customers_count)}</li>
                <li><strong>{t['am_stats']}:</strong> {t['am_stats_text'].format(with_am=customers_with_am_count, without_am=customers_without_am_count)}</li>
              </ul>
              <p style="margin-top:15px;"><strong>{t['challenges_opps']}:</strong></p>
              <p>{challenges_text}</p>
            </div>

            <h2 style="margin:25px 0 15px; font-size:18px; font-weight:600; color:#004289;">&#128200; {t['monthly_evolution']}</h2>
            <table class="data-table" role="presentation" cellpadding="0" cellspacing="0" style="width:100%; border-collapse:collapse; border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:14px; margin:20px 0; mso-table-lspace:0pt; mso-table-rspace:0pt; mso-margin-top-alt:20px; mso-margin-bottom-alt:20px;">
              <thead>
                <tr>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:left; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['month']}</th>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:right; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['orders']}</th>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:right; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['revenue']}</th>
                </tr>
              </thead>
              <tbody>"""
            
            # Monthly Rows
            for (yr, mth), row in monthly_stats.iterrows():
                # Only show recent months (e.g. last 4)
                if yr == selected_year and mth >= min(selected_months) - 2:
                    m_name = MONTH_NAMES[language][mth]
                    html_content += f"""
                    <tr>
                      <td style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px;">{m_name} {yr}</td>
                      <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{row['count']}</td>
                      <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{fmt_eur(row['sum'])}</td>
                    </tr>"""

            html_content += f"""
              </tbody>
            </table>

            <h2 style="margin:25px 0 15px; font-size:18px; font-weight:600; color:#004289;">&#128101; {t['am_revenue'].format(period=period_str)}</h2>
            <table class="data-table" role="presentation" cellpadding="0" cellspacing="0" style="width:100%; border-collapse:collapse; border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:14px; margin:20px 0; mso-table-lspace:0pt; mso-table-rspace:0pt; mso-margin-top-alt:20px; mso-margin-bottom-alt:20px;">
              <thead>
                <tr>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:left; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">Accountmanager</th>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:right; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['orders']}</th>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:right; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['revenue']}</th>
                </tr>
              </thead>
              <tbody>"""
            
            for _, row in am_stats.iterrows():
                html_content += f"""
                <tr>
                  <td style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px;">{row['am']}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{row['count']}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{fmt_eur(row['sum'])}</td>
                </tr>"""

            html_content += f"""
              </tbody>
            </table>

            <h2 style="margin:25px 0 15px; font-size:18px; font-weight:600; color:#004289;">&#128257; {t['top_returning'].format(period=period_str)}</h2>
            <table class="data-table" role="presentation" cellpadding="0" cellspacing="0" style="width:100%; border-collapse:collapse; border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:14px; margin:20px 0; mso-table-lspace:0pt; mso-table-rspace:0pt; mso-margin-top-alt:20px; mso-margin-bottom-alt:20px;">
              <thead>
                <tr>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:left; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['customer']}</th>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:right; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['orders']}</th>
                </tr>
              </thead>
              <tbody>"""
            
            for email, count in cust_freq.items():
                html_content += f"""
                <tr>
                  <td style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px;">{email}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{count}</td>
                </tr>"""

            html_content += f"""
              </tbody>
            </table>

            <h2 style="margin:25px 0 15px; font-size:18px; font-weight:600; color:#004289;">&#128181; {t['top_revenue'].format(period=period_str)}</h2>
            <table class="data-table" role="presentation" cellpadding="0" cellspacing="0" style="width:100%; border-collapse:collapse; border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:14px; margin:20px 0; mso-table-lspace:0pt; mso-table-rspace:0pt; mso-margin-top-alt:20px; mso-margin-bottom-alt:20px;">
              <thead>
                <tr>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:left; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['customer']}</th>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:right; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['revenue']}</th>
                </tr>
              </thead>
              <tbody>"""
            
            for email, rev in cust_rev.items():
                html_content += f"""
                <tr>
                  <td style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px;">{email}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{fmt_eur(rev)}</td>
                </tr>"""

            html_content += f"""
              </tbody>
            </table>

            <h2 style="margin:25px 0 15px; font-size:18px; font-weight:600; color:#004289;">&#128230; {t['top_products'].format(period=period_str)}</h2>
            <table class="data-table" role="presentation" cellpadding="0" cellspacing="0" style="width:100%; border-collapse:collapse; border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:14px; margin:20px 0; mso-table-lspace:0pt; mso-table-rspace:0pt; mso-margin-top-alt:20px; mso-margin-bottom-alt:20px;">
              <thead>
                <tr>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:left; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['product']}</th>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:right; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['quantity']}</th>
                </tr>
              </thead>
              <tbody>"""
            
            for prod, qty in prod_stats.items():
                html_content += f"""
                <tr>
                  <td style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px;">{prod}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{qty}</td>
                </tr>"""

            html_content += f"""
              </tbody>
            </table>

            <h2 style="margin:25px 0 15px; font-size:18px; font-weight:600; color:#004289;">&#128200; {t['customer_insights']}</h2>
            <table class="data-table" role="presentation" cellpadding="0" cellspacing="0" style="width:100%; border-collapse:collapse; border:1px solid #e6e8ee; font-family:'Segoe UI', Tahoma, sans-serif; font-size:14px; margin-bottom: 25px; mso-table-lspace:0pt; mso-table-rspace:0pt;">
              <thead>
                <tr>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:left; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['category']}</th>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:right; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['quantity']}</th>
                  <th style="background-color:#004289; color:#ffffff; padding:12px; text-align:right; font-weight:700; font-size:13px; text-transform:uppercase; letter-spacing:0.4px;">{t['percentage']}</th>
                </tr>
              </thead>
              <tbody>
                <tr>
                  <td style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px;">{t['total_customers']}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{total_customers_count}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">100%</td>
                </tr>
                <tr>
                  <td style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px;">{t['active_customers']}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{active_customers_count}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{f"{active_customers_count/total_customers_count*100:.1f}%" if total_customers_count else "0%"}</td>
                </tr>
                <tr>
                  <td style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px;">{t['inactive_customers']}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{inactive_customers_count}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{f"{inactive_customers_count/total_customers_count*100:.1f}%" if total_customers_count else "0%"}</td>
                </tr>
                <tr>
                  <td style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px;">{t['customers_with_am']}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{customers_with_am_count}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{f"{customers_with_am_count/total_customers_count*100:.1f}%" if total_customers_count else "0%"}</td>
                </tr>
                <tr>
                  <td style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px;">{t['customers_without_am']}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{customers_without_am_count}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{f"{customers_without_am_count/total_customers_count*100:.1f}%" if total_customers_count else "0%"}</td>
                </tr>
                <tr>
                  <td style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px;">{t['new_registrations'].format(period=period_str)}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{new_customers_count}</td>
                  <td class="number" style="padding:12px; border-bottom:1px solid #e6e8ee; font-size:14px; text-align:right; font-weight:600;">{f"{new_customers_count/total_customers_count*100:.1f}%" if total_customers_count else "0%"}</td>
                </tr>
              </tbody>
            </table>

            <div style="margin-top: 30px; padding-top: 20px; border-top: 1px solid #e6e8ee; font-size: 14px; color: #586375;">
              <p>{t['footer_contact']}</p>
            </div>
                </td>
              </tr>
            </table>
          </div>
        </td>
      </tr>
      <tr>
        <td class="footer" style="background-color:#142840; color:#ffffff; padding:35px 0;">
          <div class="footer-inner" style="width:600px; margin:0 auto; text-align:center;">
            <p style="margin:0 0 12px; font-size:14px; color:#ffffff;">&copy; {selected_year} {t['footer_rights']}</p>
            <p style="margin:0; font-size:14px;"><a href="https://kingspanstore.be" style="color:#c09a5d; text-decoration:none;">kingspanstore.be</a> | <a href="https://www.kingspanlightandair.be" style="color:#c09a5d; text-decoration:none;">www.kingspanlightandair.be</a></p>
          </div>
        </td>
      </tr>
    </table>
  </div>
</body>
</html>"""
            
            st.download_button(
                label="Download HTML Report",
                data=html_content,
                file_name=f"Sales_Report_{period_str.replace(' ', '_')}.html",
                mime="text/html"
            )
            
            st.success("Report Generated Successfully! Click above to download.")
            
            # Preview
            st.components.v1.html(html_content, height=800, scrolling=True)

    except Exception as e:
        st.error(f"An error occurred: {e}")
        import traceback
        st.text(traceback.format_exc())
else:
    st.info("Please upload both Excel files and select a period to begin.")
