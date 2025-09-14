# carpe_alfa
## LAYOUT
Huvudfönster
Titelrad: Carpe Tempus

Överst till vänster:
Textfält: "Sök artikel:" (inmatningsfält bredvid texten)

Överst till höger:
Textfält: "Ny artikel:" (inmatningsfält bredvid texten)
Knapp: "Skapa artikel"
Knapp: "Uppdatera artiklar"

Vänsterpanel
Rubrik: Tillgängliga artiklar:
Under rubriken: Lista (vertikal lista med artikelnummer och storlek, t.ex. "BR101046 (46x43)").

Mittenpanel (övre)
Rubrik: Välj operation:
Dropdown med operationer (ex. "1 - Tillskärning av Tyg").

Till höger om dropdown:
Rubrik: Extra operatörer:
Dropdown med siffror (0–...).

Under dropdown-menyerna: Fyra stora knappar på rad:
Blå: "Starta ställtid"
Grön: "Starta produktion"
Röd: "Avsluta produktion"
Orange: "Avvikelse"

Under knapparna:
Text: "Användare: [namn]"
Text: "Pågående jobb: ..." (visar artikel, operation och tid)
Text: "App-sessionstid: ..."
Text: "Navision-sessionstid: ..."

Mittenpanel (höger, status)
Två statusfält med färgindikatorer:
"Status:" (grön prick bredvid)
"Åtkomst:" (röd prick bredvid)

Nedre panel
Rubrik: Artikeldetaljer:
Tabell med kolumner:
Datum
Användare
Artikel
Operation
Antal
Ställtid
Produktionstid
Typ av jobb
Skrotade
Tid/del (min)
Tabellen visar rader med data (tidigare jobb).

Sidfot
Vänster: Text: "© 2025, Emanuel Teljemo, Alla rättigheter förbehållna."
Mitten: "Version: 1.0"
Höger: Knapp: "Kontrollera Uppdatering"

### Layoutdelar i din kod

I din fil carpe_tempus_1_01.txt börjar själva gränssnittet byggas i sektionen:

# ---------------------------------------------------------------------------------
# [HUVUD-UI OCH PROGRAMSTART]
# ---------------------------------------------------------------------------------

### Viktiga delar:

1. Huvudfönster

root = tk.Tk()
root.title("Carpe Tempus")
root.geometry("1150x700")


2. Input-rad överst (sök, ny artikel, knappar)

input_frame = tk.Frame(main_frame)
...
tk.Label(input_frame, text="Sök artikel:")
tk.Entry(input_frame, textvariable=search_var)
tk.Label(input_frame, text="Ny artikel:")
tk.Entry(input_frame, textvariable=new_article_var)
tk.Button(input_frame, text="Skapa artikel")
tk.Button(input_frame, text="Uppdatera artiklar")


3. Split-panel (vänster = artiklar, höger = operation + knappar + status)

split_frame = tk.Frame(main_frame)
• Vänster panel
article_frame = tk.Frame(split_frame)
article_selection_tree = ttk.Treeview(article_selection_frame, columns=("Article",))
• Höger panel
right_frame = tk.Frame(split_frame)


4. Operation och extra operatörer

operation_combobox = ttk.Combobox(operation_frame, textvariable=selected_operation)
operators_spinbox = ttk.Spinbox(operators_frame, from_=0, to=10, textvariable=selected_operators)


5. Knapparna (ställtid, start, stop, avvikelse)

setup_button = tk.Button(button_frame, text="Starta ställtid", bg="blue", fg="white")
start_button = tk.Button(button_frame, text="Starta produktion", bg="green", fg="white")
stop_button = tk.Button(button_frame, text="Avsluta produktion", bg="red", fg="black")
deviation_button = tk.Button(button_frame, text="Avvikelse", bg="orange", fg="black")


6. Status-indikatorer (gröna/röda prickar)

access_indicator = tk.Label(button_frame, text="●", font=("Arial", 20), fg="red")
drive_indicator = tk.Label(button_frame, text="●", font=("Arial", 20), fg="red")


7. Pågående jobb-ruta

running_frame = tk.Frame(right_frame)


8. Artikeldetaljer-tabell

details_tree = ttk.Treeview(main_frame, columns=(...), show="headings")


9. Sidfot (info, version, uppdateringsknapp)

info_frame = tk.Frame(root)
made_by_label = tk.Label(info_frame, text="© 2025, Emanuel Teljemo...")
tk.Label(info_frame, text=f"Version: {VERSION}")
