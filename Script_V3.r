#####-----------------------IMPORTATION DES LIBRAIRIES-------------###########
library(kableExtra)
library(officer)
library(flextable)
library(openxlsx)
library(scales)
library(purrr)
library(xts)
library(lubridate)
library(dplyr)
library(tidyr)
library(ggplot2)
library(rvg)
library(patchwork)
library(zip)
library(ggrepel)
library(readxl)
library(readr)
library(stringi)
library(data.table)
library(DBI)
library(odbc)

  
cat("###############################################\n")
cat("######### P3 - FIABILISATION DES DONNEES KYC \n")
cat("######### ETAPE 1/3: CHAMPS A FIABILISER\n")
cat("##############################################\n")



#####-----------------------####-------------###########
chemin="C://Fiabilisation KYC//R//"
chemin2="C://KYC_Filiales//"
chemin_python="C://Fiabilisation KYC//Python//data//"
chemin_7z <- '"C://Program Files//7-Zip//7z.exe"'
#####-----------------------####-------------###########



## ID pour l'envoi via smtp

id=read.csv2(paste0(chemin,"id_sfkyc.csv"))

from <- id[1,1]



#####-----------------------####-------------###########

#####-------- LECTURE DE prerequis.csv --------###########
# Le fichier contient un bloc d'entete libre (ligne "Trimestre", lignes vides)
# puis une ligne de titres de colonnes suivie d'une ligne par filiale :
#
#   Trimestre;1;;;;;
#   ;;;;;;
#   Filiales;y/n;Devise;Flux(mois/jour);pm_exclure;pp_exclure;rc_exclure
#   SN;y;XOF;jour;EIN,ORG;;
#
# On repere la ligne de titres, on l'utilise comme noms de colonnes et on ne
# garde que les lignes situees en dessous. Les colonnes sont ensuite adressees
# PAR NOM (prerequis_col) : ajouter, deplacer ou renommer une colonne dans le
# CSV ne casse plus le script.
prerequis_brut <- read.csv2(paste(sep="",chemin,"prerequis.csv"),
                            header = FALSE, stringsAsFactors = FALSE,
                            colClasses = "character")

# "Flux(mois/jour)" -> "flux_mois_jour" ; "y/n" -> "y_n"
normaliser_nom <- function(x) {
  x <- tolower(trimws(as.character(x)))
  x <- gsub("[^a-z0-9]+", "_", x)
  gsub("^_+|_+$", "", x)
}

# Ligne de titres = premiere ligne dont la 1re cellule vaut "filiale(s)".
# Repli : premiere ligne contenant une cellule "pm_exclure".
ligne_titres <- {
  c1 <- normaliser_nom(prerequis_brut[[1]])
  i  <- which(c1 %in% c("filiale", "filiales"))
  if (length(i) == 0) {
    i <- which(apply(prerequis_brut, 1,
                     function(l) any(normaliser_nom(l) == "pm_exclure")))
  }
  if (length(i) == 0) NA_integer_ else i[1]
}
if (is.na(ligne_titres)) {
  stop("prerequis.csv : ligne de titres introuvable (attendue : 'Filiales;y/n;Devise;...').")
}

# Trimestre : valeur en 2e colonne de la ligne "Trimestre" (cherchee dans le
# bloc d'entete), avec repli sur l'ancien emplacement (1re ligne, 2e colonne).
trimestre_actuel <- {
  c1 <- normaliser_nom(prerequis_brut[[1]])
  i  <- which(c1 == "trimestre")
  as.character(prerequis_brut[[2]][if (length(i)) i[1] else 1])
}

prerequis <- prerequis_brut[seq.int(ligne_titres + 1, nrow(prerequis_brut)), , drop = FALSE]
noms_prerequis <- normaliser_nom(unlist(prerequis_brut[ligne_titres, ], use.names = FALSE))
noms_prerequis[noms_prerequis == ""] <- paste0("col_", which(noms_prerequis == ""))
colnames(prerequis) <- noms_prerequis
rownames(prerequis) <- NULL

# Alias historiques : le reste du script manipule prerequis$infos / $argument.
prerequis$infos    <- trimws(as.character(prerequis[[1]]))
prerequis$argument <- trimws(as.character(prerequis[[2]]))
prerequis <- prerequis[prerequis$infos != "", , drop = FALSE]

# Indice de la premiere colonne portant l'un des noms donnes (NA si aucune).
prerequis_col <- function(...) {
  noms <- normaliser_nom(c(...))
  i <- which(colnames(prerequis) %in% noms)
  if (length(i) == 0) NA_integer_ else i[1]
}

# Valeur brute d'une colonne pour une filiale ("" si colonne ou ligne absente).
prerequis_valeur <- function(fil, ...) {
  j <- prerequis_col(...)
  if (is.na(j)) return("")
  v <- prerequis[[j]][prerequis$infos == fil]
  if (length(v) == 0 || is.na(v[1])) "" else trimws(as.character(v[1]))
}


filiale=prerequis$infos[prerequis$argument=="y"]


#####-------- PERIODE DE TRAITEMENT PAR FILIALE --------###########
# La periode flux est definie par filiale via une colonne "PERIODE" de
# prerequis.csv (sur la ligne de la filiale) :
#   "jour" (ou colonne vide / absente) -> la derniere journee presente dans les
#                                         donnees de la filiale
#   "mois"                             -> le mois calendaire precedant le mois
#                                         en cours (base sur Sys.Date())
# La valeur est cherchee dans toute la ligne, donc la position de la colonne
# dans le CSV n'a pas d'importance.
periode_mode <- function(fil) {
  # Colonne "Flux(mois/jour)" (alias PERIODE) de prerequis.csv.
  v <- tolower(prerequis_valeur(fil, "flux_mois_jour", "flux", "periode"))
  if (v %in% c("mois", "mois_precedent", "m")) return("mois")
  if (v %in% c("jour", "j")) return("jour")

  # Repli : colonne absente ou vide -> on cherche la valeur n'importe ou sur la
  # ligne de la filiale (ancien comportement).
  ligne <- prerequis[prerequis$infos == fil, , drop = FALSE]
  if (nrow(ligne) == 0) return("jour")
  vals <- tolower(trimws(as.character(unlist(ligne, use.names = FALSE))))
  if (any(vals %in% c("mois", "mois_precedent", "m"), na.rm = TRUE)) return("mois")
  "jour"
}


#####-------- CODES EXCLUS (pm_exclure / pp_exclure / rc_exclure) --------#####
# La colonne "pm_exclure" de prerequis.csv porte, sur la ligne de la filiale,
# les codes AGEC a exclure de la liste des PM sans RCSNO, separes par une
# virgule. Parentheses, espaces et casse sont tolerees : "EIN,ORG",
# "(EIN, ORG)" et "ein ; org" donnent le meme resultat.
# Colonne vide ou absente -> repli sur PM_EXCLURE_DEFAUT.
#PM_EXCLURE_DEFAUT <- c("ORG", "ISB")
PM_EXCLURE_DEFAUT <- c()

# "(EIN, ORG)" -> c("EIN","ORG") ; "" / NA -> character(0)
parse_codes_agec <- function(x) {
  if (length(x) == 0 || is.na(x[1])) return(character(0))
  codes <- toupper(trimws(unlist(strsplit(gsub("[()]", "", as.character(x[1])), "[,;]"))))
  codes[codes != "" & !is.na(codes)]
}

# Codes d'une colonne "*_exclure" pour une filiale.
codes_exclure <- function(fil, colonne, defaut = character(0)) {
  codes <- parse_codes_agec(prerequis_valeur(fil, colonne))
  if (length(codes) == 0) defaut else codes
}

pm_agec_exclus <- function(fil) codes_exclure(fil, "pm_exclure", PM_EXCLURE_DEFAUT)

# Bornes de la periode.
#   premier_jour = debut INCLUS
#   jour_recent  = borne haute EXCLUE
# Les filtres flux utilisent DATOUV >= premier_jour & DATOUV < jour_recent.
#
#   mode "mois" : premier_jour = 1er jour du mois calendaire precedent,
#                 dernier jour couvert = dernier jour de ce meme mois.
#                 Libelle rapport : la periode complete "du .. au ..".
#   mode "jour" : premier_jour = DATOUV la plus recente de pp_stock ET pm_stock,
#                 jour_recent  = le lendemain (la periode ne couvre que ce jour).
#                 Libelle rapport : cette seule date.
#
# dates_dispo = ensemble des DATOUV disponibles (PP et PM reunies).
bornes_periode <- function(fil, dates_dispo) {
  mode <- periode_mode(fil)
  if (mode == "mois") {
    fin   <- floor_date(Sys.Date(), "month")
    debut <- fin %m-% months(1)
  } else {
    # Les fichiers sources contiennent des DATOUV aberrantes (saisies erronees,
    # annees futures type 2090). On ignore tout ce qui est posterieur a
    # aujourd'hui, sinon la periode entiere est calee sur la date fantaisiste.
    dates_ok <- dates_dispo[!is.na(dates_dispo) & dates_dispo <= Sys.Date()]
    if (length(dates_ok) == 0) {
      stop("Aucune date exploitable (<= aujourd'hui) dans les donnees de ", fil)
    }
    ignorees <- sum(!is.na(dates_dispo) & dates_dispo > Sys.Date())
    if (ignorees > 0) {
      cat("######### ATTENTION", fil, ":", ignorees,
          "date(s) DATOUV posterieure(s) a aujourd'hui ignoree(s), max =",
          format(max(dates_dispo, na.rm = TRUE), "%d/%m/%Y"), "\n")
    }
    derniere <- max(dates_ok)
    debut <- derniere
    fin   <- derniere + 1
  }
  # Libelle affiche dans le rapport : periode complete en mode mois,
  # date unique en mode jour.
  libelle <- if (mode == "mois") {
    paste0("du ", format(debut, "%d/%m/%Y"), " au ", format(fin - 1, "%d/%m/%Y"))
  } else {
    paste0("du ", format(fin - 1, "%d/%m/%Y"))
  }
  cat("######### Periode retenue pour", fil, "(mode", mode, ") :",
      libelle, "\n")
  list(debut = debut, fin = fin, mode = mode, libelle = libelle)
}


inc=c("")


pp_fields <- c("PAYNAIS","PROFESSION","SALAIRE","NUMID","CODAPE","TEL","DATNAIS","ADRESSE","DATVALID","ORIGINE_REVENU","PAYS_RESID")
pp_labels <- c("Lieu de Naissance","Profession","Salaire","NIN","Code agent économique","Téléphone","Adresse","Date de naissance","Date de validité CIN","Origine du revenu", "Pays de résidence")

pm_fields <- c("CODAPE","AGEC","CAPITAL","CA","RESULTAT","RCSNO","ORIGINE_REVENU","TEL")
pm_labels <- c("Secteur d'activité","Code agent économique","Capital social","Chiffre d'affaires","Résultat net","Numéro de registre de commerce","Origine du revenu","Téléphone")




filiale <- prerequis$infos[prerequis$argument == "y"]

# Fonctions utilitaires de nettoyage
is_rep <- function(x) {
  (grepl("XX", x) & nchar(x)==2) | (grepl("RAS", x) & nchar(x)==3) | (grepl("R.A.S.", x) & nchar(x)==6) | (grepl("R.A.S", x) & nchar(x)==5) | nchar(x)==1
}

clean_and_mark_anomalies <- function(df, start_col = 4) {
  if (is.null(df) || ncol(df) < start_col) return(df)
  df[] <- lapply(df, function(x) if (is.character(x)) trimws(x) else x)
  cols <- start_col:ncol(df)
  df[cols] <- Map(function(x, nm) {
    x[is.na(x)] <- ""
    rep_idx <- is_rep(x)
    if (any(rep_idx, na.rm = TRUE)) x[rep_idx] <- paste("ANOMALIE", nm)
    x
  }, df[cols], names(df)[cols])
  df
}

# Fonctions de calcul des taux
to_pct <- function(x) {
  x_num <- suppressWarnings(as.numeric(x))
  ifelse(is.na(x_num), "N/A", paste0(floor(x_num), "%"))
}

# Taux de complétude tronqué à l'entier (jamais arrondi). 100 % uniquement si
# aucun défaut (n_vide == 0) ; sinon plafonné à 99 % pour ne jamais afficher un
# 100 % trompeur (protège aussi d'un arrondi flottant à 100). Même logique que
# _rate_floor_1dec côté Django.
rate_floor_no100 <- function(n_vide, total) {
  if (is.null(total) || length(total) == 0 || is.na(total) || total <= 0) return(NA_real_)
  if (is.na(n_vide) || n_vide <= 0) return(100)
  min(floor(100 * (1 - (n_vide / total))), 99)
}

champ_rate <- function(df, champ) {
  if (is.null(df) || nrow(df) == 0 || !(champ %in% colnames(df))) return(NA_real_)
  n_vide <- nrow(df[trimws(as.character(df[[champ]])) == "" | is.na(df[[champ]]), , drop = FALSE])
  rate_floor_no100(n_vide, nrow(df))
}

global_rate <- function(df, fields) {
  if (is.null(df) || nrow(df) == 0) return(NA_real_)
  total_cells <- length(fields) * nrow(df)
  n_empty <- sum(sapply(df[, intersect(fields, colnames(df)), drop=FALSE], function(x) sum(trimws(as.character(x)) == "" | is.na(x))))
  rate_floor_no100(n_empty, total_cells)
}

client_completion_rate <- function(df, fields) {
  if (is.null(df) || nrow(df) == 0) return(NA_real_)
  block <- df[, intersect(fields, colnames(df)), drop = FALSE]
  row_has_empty <- Reduce(`|`, lapply(block, function(x) trimws(as.character(x)) == "" | is.na(x)))
  rate_floor_no100(sum(row_has_empty), nrow(df))
}

# --- FONCTION DE STYLE (RENDU EXACT DE L'IMAGE) ---
style_ft_exact <- function(df) {
  ft <- flextable(df)
  
  # Fusion verticale de la premià¨re colonne (Libellés de groupe)
  ft <- merge_v(ft, j = 1) %>% valign(j = 1, valign = "center")
  
  # En-tàªte : bleu marine, texte blanc, gras (palette "exemple rapport")
  ft <- bg(ft, part = "header", bg = "#1B2A4A") %>%
        color(part = "header", color = "white") %>%
        bold(part = "header", bold = TRUE)
  ft <- font(ft, fontname = "Calibri", part = "all")
  ft <- color(ft, color = "#333333", part = "body")
  # Colonne de regroupement : fond vert tres clair
  ft <- bg(ft, j = 1, bg = "#F0F7F4", part = "body")
  ft <- color(ft, j = 1, color = "#1B2A4A", part = "body")
  ft <- bold(ft, j = 1, bold = TRUE, part = "body")

  # Alignements
  ft <- align(ft, align = "center", part = "header")
  ft <- align(ft, j = 1:2, align = "left", part = "body")
  ft <- align(ft, j = 3:4, align = "center", part = "body")
  
  # Lignes de synthà¨se du bas : Gras + Alignement spécifique
  n_rows <- nrow(df)
  ft <- bold(ft, i = (n_rows-1):n_rows, bold = TRUE)
  ft <- align(ft, i = (n_rows-1):n_rows, j = 2, align = "right") # Aligne "Approche par..." à  droite
  
  # Lignes de synthese : fond vert clair
  ft <- bg(ft, i = (n_rows-1):n_rows, bg = "#C8E6C9", part = "body")
  ft <- color(ft, i = (n_rows-1):n_rows, color = "#1B2A4A", part = "body")

  # Hors seuil : rouge doux sur fond rose (au lieu de l'orange)
  ft <- bg(ft, j = "Flux (3)",
           i = ~ as.numeric(gsub("%", "", `Flux (3)`)) < 100,
           bg = "#FDE2E1", part = "body")
  ft <- color(ft, j = "Flux (3)",
              i = ~ as.numeric(gsub("%", "", `Flux (3)`)) < 100,
              color = "#C0392B", part = "body")

  ft <- bg(ft, j = "Stock (4)",
           i = ~ as.numeric(gsub("%", "", `Stock (4)`)) < 90,
           bg = "#FDE2E1", part = "body")
  ft <- color(ft, j = "Stock (4)",
              i = ~ as.numeric(gsub("%", "", `Stock (4)`)) < 90,
              color = "#C0392B", part = "body")

  ft <- bold(ft, j = c("Flux (3)", "Stock (4)"), bold = TRUE, part = "body")

  # Filets gris clairs uniquement (plus de bordures noires)
  ft <- border_remove(ft)
  ft <- hline(ft, border = fp_border(color = "#E0E0E0", width = 0.75), part = "body")
  ft <- vline(ft, border = fp_border(color = "#E0E0E0", width = 0.75), part = "all")
  ft <- border_outer(ft, border = fp_border(color = "#E0E0E0", width = 0.75))

  # Tailles et dimensions compactes pour une slide 10 x 5.625 pouces
  ft <- fontsize(ft, size = 9, part = "all")
  ft <- padding(ft, padding.top = 1, padding.bottom = 1, padding.left = 2, padding.right = 2, part = "all")
  ft <- height_all(ft, height = 0.24, part = "all")
  ft <- width(ft, j = 1, width = 2.1)
  ft <- width(ft, j = 2, width = 4.25)
  ft <- width(ft, j = 3, width = 1.5)
  ft <- width(ft, j = 4, width = 1.5)
  
  return(ft)
}

agent_completion_rates <- function(df, fields) {
  if (is.null(df) || nrow(df) == 0 || !("EXPL" %in% colnames(df))) return(numeric(0))
  agents <- sort(unique(trimws(as.character(df$EXPL))))
  agents <- agents[agents != "" & !is.na(agents)]
  rates <- sapply(agents, function(agent) {
    global_rate(df[trimws(as.character(df$EXPL)) == agent, , drop = FALSE], fields)
  })
  as.numeric(rates[!is.na(rates)])
}

agent_distribution_row <- function(rates) {
  labels <- c("[0, 50%]", "]50% , 60%]", "]60% , 80%]", "]80% , 100%]")
  if (length(rates) == 0) return(setNames(rep("0%", length(labels)), labels))
  cuts <- cut(rates, breaks = c(0, 50, 60, 80, 100), include.lowest = TRUE, right = TRUE, labels = labels)
  counts <- table(factor(cuts, levels = labels))
  setNames(paste0(floor(100 * as.numeric(counts) / length(rates)), "%"), labels)
}

style_slide5_table <- function(df, widths = NULL) {
  ft <- flextable(df)
  ft <- bg(ft, part = "header", bg = "#1B2A4A") %>%
    color(part = "header", color = "white") %>%
    bold(part = "header", bold = TRUE)
  ft <- font(ft, fontname = "Calibri", part = "all")
  ft <- color(ft, color = "#333333", part = "body")
  ft <- fontsize(ft, size = 9, part = "all")
  ft <- padding(ft, padding.top = 3, padding.bottom = 3, padding.left = 4, padding.right = 4, part = "all")
  ft <- align(ft, align = "center", part = "all")
  ft <- align(ft, j = 1, align = "left", part = "body")
  ft <- bold(ft, j = 1, bold = TRUE, part = "body")
  # lignes alternees, filets gris clairs uniquement
  if (nrow(df) > 1) {
    pairs <- seq(2, nrow(df), by = 2)
    if (length(pairs) > 0) ft <- bg(ft, i = pairs, bg = "#F0F7F4", part = "body")
  }
  ft <- border_remove(ft)
  ft <- hline(ft, border = fp_border(color = "#E0E0E0", width = 0.75), part = "body")
  ft <- border_outer(ft, border = fp_border(color = "#E0E0E0", width = 0.75))
  if (!is.null(widths)) {
    for (j in seq_along(widths)) ft <- width(ft, j = j, width = widths[j])
  }
  ft
}


#####--------- CONSTRUCTION DU RAPPORT DG ---------#####
# Charte reprise a l'identique de "exemple rapport.pptx" : couleurs, police,
# geometrie des bandeaux, cartouches et pied de page ont ete releves dans le
# XML de ce fichier et sont reproduits tels quels ci-dessous.

DG_VERT   <- "#009A56"   # vert BOA : bandeau de titre, pastilles, libelles
DG_VERT_C <- "#C8E6C9"   # vert clair : lignes de synthese
DG_NAVY   <- "#1B2A4A"   # bleu marine : texte fort, barres de cartouche
DG_PANEL  <- "#F5F5F5"   # fond des cartouches
DG_FOND   <- "#F0F7F4"   # fond vert tres pale
DG_GRIS   <- "#E0E0E0"   # filets, bande de pied de page
DG_TXT    <- "#333333"   # texte courant
DG_MUET   <- "#888888"   # texte de pied de page
DG_FONT   <- "Calibri"

# Geometrie relevee dans "exemple rapport.pptx" (slide 10 x 5,625 po)
# Palette des cartes KPI : tons sobres de facture bancaire (vert BOA, marine,
# bleus ardoise, gris-bleu). DG_ALERTE, un brique attenue, remplace la couleur
# d'origine quand un seuil n'est pas tenu.
DG_FUN    <- c("#009A56", "#1B2A4A", "#2E7D8F", "#3D5A80", "#4E6E5D", "#5B6B7B")
DG_ALERTE <- "#B03A2E"

DG_HDR_H    <- 0.62      # hauteur du bandeau vert de titre
DG_MARGE    <- 0.40      # marge de contenu
DG_RULE_Y   <- 1.08      # filet vert sous le libelle de section
DG_FOOT_Y   <- 5.35      # bande grise de pied de page
DG_FOOT_H   <- 0.28

# Base de presentation : on repart de "exemple rapport.pptx" pour heriter
# exactement de son theme, de son masque, de ses polices et de son format 16/9,
# en supprimant ses slides. Si le fichier est absent, on retombe sur un
# template 16/9 genere a partir de celui d'officer.
.dg_tpl169 <- NULL

# Maquette prete a remplir : en-tetes, pieds de page, logo, sommaire et pages
# de garde y sont deja composes. Le script n'y depose que le contenu.
dg_maquette_path <- function() {
  for (f in c("exemple_rapport_kyc.pptx", "exemple rapport kyc.pptx",
              "exemple_rapport_KYC.pptx")) {
    p <- paste0(chemin, f)
    if (file.exists(p)) return(p)
  }
  NULL
}

dg_base_path <- function() {
  for (f in c("exemple rapport.pptx", "exemple_rapport.pptx")) {
    p <- paste0(chemin, f)
    if (file.exists(p)) return(p)
  }
  NULL
}

dg_pptx <- function() {
  # 1) maquette prete a remplir : on garde ses slides telles quelles
  maq <- dg_maquette_path()
  if (!is.null(maq)) {
    p <- tryCatch(read_pptx(maq), error = function(e) NULL)
    if (!is.null(p)) {
      if (length(p) < 7) {
        cat("######### ATTENTION : la maquette ne contient que", length(p),
            "slides (7 attendues)\n")
      }
      return(p)
    }
    cat("######### ATTENTION : maquette 'exemple_rapport_kyc.pptx' illisible\n")
  } else {
    cat("######### ATTENTION : 'exemple_rapport_kyc.pptx' introuvable dans",
        chemin, "-> deck reconstruit\n")
  }
  # 2) repli : theme de "exemple rapport.pptx", slides supprimees
  base <- dg_base_path()
  if (!is.null(base)) {
    p <- tryCatch({
      x <- read_pptx(base)
      while (length(x) > 0) x <- remove_slide(x, 1)
      x
    }, error = function(e) NULL)
    if (!is.null(p)) return(p)
    cat("######### ATTENTION : 'exemple rapport.pptx' illisible, theme par defaut utilise\n")
  } else {
    cat("######### ATTENTION : 'exemple rapport.pptx' introuvable dans", chemin,
        "-> theme par defaut\n")
  }
  dg_pptx_fallback()
}

dg_pptx_fallback <- function() {
  if (is.null(.dg_tpl169) || !file.exists(.dg_tpl169)) {
    src <- system.file(package = "officer", "template", "template.pptx")
    ok <- src != "" && requireNamespace("zip", quietly = TRUE)
    if (ok) {
      d <- file.path(tempdir(), "dg_tpl169_src")
      unlink(d, recursive = TRUE); dir.create(d, recursive = TRUE)
      utils::unzip(src, exdir = d)
      p <- file.path(d, "ppt", "presentation.xml")
      x <- paste(readLines(p, warn = FALSE), collapse = "")
      x <- sub('sldSz[^/]*?/>', 'sldSz cx="9144000" cy="5143500" type="screen16x9"/>', x)
      con <- file(p, open = "wb"); writeBin(charToRaw(x), con); close(con)
      out <- file.path(tempdir(), "dg_template_169.pptx"); unlink(out)
      zip::zip(zipfile = out, files = list.files(d, recursive = TRUE, all.files = TRUE),
               root = d, mode = "mirror")
      .dg_tpl169 <<- out
    }
  }
  if (!is.null(.dg_tpl169) && file.exists(.dg_tpl169)) read_pptx(.dg_tpl169) else read_pptx()
}

# Layout sans placeholder : on positionne tout en coordonnees absolues.
dg_layout <- function(ppt) {
  ls <- layout_summary(ppt)
  i <- which(ls$layout %in% c("Blank", "Vide"))[1]
  if (is.na(i)) i <- which(ls$layout %in% c("Title Only", "Titre seul"))[1]
  if (is.na(i)) i <- nrow(ls)
  list(layout = ls$layout[i], master = ls$master[i])
}

dg_add <- function(ppt) {
  lay <- dg_layout(ppt)
  add_slide(ppt, layout = lay$layout, master = lay$master)
}

`%||%` <- function(a, b) if (is.null(a)) b else a

# --- Bloc colore ---
# ATTENTION : officer n'exporte PAS fp_par(shading.color) vers PowerPoint.
# Le seul aplat de couleur reellement rendu passe par une table. Toute zone
# coloree du rapport est donc construite ici, via un flextable sans en-tete.
# `lignes` : liste de list(txt, size, bold, color, bg, align, h)
dg_bloc <- function(lignes, largeur, pad_left = 6) {
  ft <- flextable(data.frame(x = vapply(lignes, function(l) l$txt, character(1)),
                             stringsAsFactors = FALSE))
  ft <- delete_part(ft, part = "header")
  ft <- border_remove(ft)
  ft <- font(ft, fontname = DG_FONT, part = "all")
  ft <- width(ft, j = 1, width = largeur)
  for (i in seq_along(lignes)) {
    l <- lignes[[i]]
    ft <- bg(ft, i = i, bg = l$bg %||% "transparent", part = "body")
    ft <- color(ft, i = i, color = l$color %||% DG_TXT, part = "body")
    ft <- fontsize(ft, i = i, size = l$size %||% 11, part = "body")
    ft <- bold(ft, i = i, bold = isTRUE(l$bold), part = "body")
    ft <- align(ft, i = i, align = l$align %||% "left", part = "body")
    if (!is.null(l$h)) ft <- height(ft, i = i, height = l$h, part = "body")
  }
  ft <- padding(ft, padding.top = 2, padding.bottom = 2,
                padding.left = pad_left, padding.right = 6, part = "body")
  set_table_properties(ft, layout = "fixed")
}

# Bandeau plein de couleur unie (filet, fond de section)
dg_filet <- function(couleur, largeur, hauteur = 0.06) {
  dg_bloc(list(list(txt = "", bg = couleur, size = 4, h = hauteur)), largeur)
}

# Paragraphe simple (texte sans aplat : `bg` est ignore, cf. dg_bloc)
dg_par <- function(txt, size = 12, bold = FALSE, color = "black",
                   align = "left", bg = "transparent", italic = FALSE) {
  fpar(ftext(txt, fp_text(color = color, font.size = size, bold = bold,
                          italic = italic, font.family = DG_FONT)),
       fp_p = fp_par(text.align = align, shading.color = bg,
                     padding.top = 3, padding.bottom = 3,
                     padding.left = 6, padding.right = 6))
}

# --- En-tete de slide ---
# Bandeau bleu marine pleine largeur : eyebrow "BANK OF AFRICA - BMCE GROUP",
# titre de section, et sous-titre optionnel aligne a droite (date, perimetre).
# Un filet vert souligne le bandeau.
# Bandeau vert pleine largeur + cartouche blanc "BANK OF AFRICA / BMCE GROUP"
# a droite, exactement comme dans "exemple rapport.pptx".
dg_entete <- function(ppt, ppt_titre, W, sous_titre = NULL, M = DG_MARGE) {
  # bandeau vert + titre blanc 18 pt
  ppt <- ph_with(ppt, value = dg_bloc(list(
      list(txt = ppt_titre, size = 18, bold = TRUE, color = "white",
           bg = DG_VERT, h = DG_HDR_H)), W, pad_left = 22),
    location = ph_location(left = 0, top = 0, width = W, height = DG_HDR_H))
  # cartouche blanc en surimpression a droite : logo si disponible, sinon texte
  ppt <- ph_with(ppt, value = dg_bloc(list(
      list(txt = "", bg = "white", size = 4, h = 0.55)), 1.5),
    location = ph_location(left = W - 1.5, top = 0, width = 1.5, height = 0.55))
  p_logo <- dg_logo_path()
  if (!is.null(p_logo)) {
    r <- dg_logo_ratio(p_logo)
    lh <- 0.36; lw <- lh * r
    if (lw > 1.34) { lw <- 1.34; lh <- lw / r }
    ppt <- ph_with(ppt, value = external_img(p_logo, width = lw, height = lh),
                   location = ph_location(left = W - 1.5 + (1.5 - lw)/2,
                                          top = (0.55 - lh)/2, width = lw, height = lh))
  } else {
    ppt <- ph_with(ppt, value = dg_bloc(list(
        list(txt = "BANK OF AFRICA", size = 7, bold = TRUE, color = DG_NAVY,
             align = "center", bg = "white", h = 0.28),
        list(txt = "BMCE GROUP", size = 6, color = DG_NAVY,
             align = "center", bg = "white", h = 0.24)), 1.5),
      location = ph_location(left = W - 1.5, top = 0, width = 1.5, height = 0.55))
  }
  if (!is.null(sous_titre)) {
    ppt <- ph_with(ppt, value = block_list(
        dg_par(sous_titre, size = 9, bold = TRUE, color = DG_MUET, align = "right")),
      location = ph_location(left = W/2, top = DG_RULE_Y - 0.3, width = W/2 - M, height = 0.26))
  }
  ppt
}

# Libelle de section (pastille verte) + filet vert, sous le bandeau de titre
dg_section <- function(ppt, txt, W, largeur = 2.4, M = DG_MARGE) {
  ppt <- ph_with(ppt, value = dg_bloc(list(
      list(txt = txt, size = 11, bold = TRUE, color = "white",
           bg = DG_VERT, h = 0.32)), largeur, pad_left = 10),
    location = ph_location(left = M, top = 0.80, width = largeur, height = 0.32))
  ph_with(ppt, value = dg_filet(DG_VERT, W - 2*M, 0.04),
          location = ph_location(left = M, top = DG_RULE_Y, width = W - 2*M, height = 0.04))
}

# Cartouche : panneau gris clair coiffe d'une barre de titre coloree
dg_cartouche <- function(ppt, titre, left, top, largeur, hauteur,
                         couleur = DG_NAVY) {
  ppt <- ph_with(ppt, value = dg_bloc(list(
      list(txt = titre, size = 9, bold = TRUE, color = "white",
           bg = couleur, h = 0.28)), largeur, pad_left = 10),
    location = ph_location(left = left, top = top, width = largeur, height = 0.28))
  ph_with(ppt, value = dg_bloc(list(
      list(txt = "", bg = DG_PANEL, size = 4, h = hauteur - 0.28)), largeur),
    location = ph_location(left = left, top = top + 0.28,
                           width = largeur, height = hauteur - 0.28))
}

# Compatibilite : ancien nom
dg_titre <- function(ppt, txt, W, M = 0.3) dg_entete(ppt, txt, W, NULL, M)

# --- Pied de page --- bande grise pleine largeur + mention a droite
dg_pied <- function(ppt, W, H, num = NULL, M = DG_MARGE) {
  ppt <- ph_with(ppt, value = dg_filet(DG_GRIS, W, DG_FOOT_H),
                 location = ph_location(left = 0, top = DG_FOOT_Y, width = W, height = DG_FOOT_H))
  mention <- if (is.null(num)) "Conformité BOA Group  |  Pôle Projet"
             else paste0("Conformité BOA Group  |  Pôle Projet        ", num)
  ppt <- ph_with(ppt, value = block_list(
      dg_par(mention, size = 8, color = DG_MUET, align = "right")),
    location = ph_location(left = W - 3.6, top = DG_FOOT_Y + 0.02, width = 3.4, height = 0.24))
  ppt
}

# --- Carte KPI ---
# Pastille verte facon "01/02/03" de l'exemple : valeur sur aplat colore,
# libelle en bleu marine sur le panneau gris.
dg_kpi <- function(valeur, libelle, couleur = DG_VERT, largeur = 1.5) {
  dg_bloc(list(
    list(txt = valeur,  size = 20, bold = TRUE, color = "white",
         align = "center", bg = couleur, h = 0.50),
    list(txt = libelle, size = 7.5, bold = TRUE, color = DG_NAVY,
         align = "center", bg = DG_PANEL, h = 0.38)), largeur, pad_left = 3)
}

# Depose une rangee de cartes KPI reparties sur la largeur utile.
dg_kpi_row <- function(ppt, kpis, W, top, M = 0.3, hauteur = 0.9, gap = 0.1) {
  n <- length(kpis)
  if (n == 0) return(ppt)
  larg <- (W - 2*M - (n - 1) * gap) / n
  for (i in seq_len(n)) {
    k <- kpis[[i]]
    ppt <- ph_with(ppt, value = dg_kpi(k$valeur, k$libelle, k$couleur %||% DG_VERT, larg),
                   location = ph_location(left = M + (i - 1) * (larg + gap),
                                          top = top, width = larg, height = hauteur))
  }
  ppt
}

# Logo de la banque (depose dans `chemin`). Absent -> slide generee sans logo.
dg_logo_path <- function() {
  for (f in c("logo_PNG.png", "logo_PNG.PNG", "logo_PNG", "logo_PNG.jpg",
              "logo_BOA.png", "logo.png")) {
    p <- paste0(chemin, f)
    if (file.exists(p)) return(p)
  }
  NULL
}

# Largeur/hauteur reelle du logo, pour ne pas le deformer.
# Lecture directe de l'en-tete PNG (octets 17-24) : exacte et sans dependance.
# Repli sur le package png (JPEG ou en-tete illisible), puis sur 3.
dg_logo_ratio <- function(p) {
  r <- tryCatch({
    con <- file(p, "rb"); on.exit(close(con), add = TRUE)
    h <- readBin(con, "raw", n = 24)
    if (length(h) == 24 && identical(as.integer(h[2:4]), c(80L, 78L, 71L))) {
      w  <- sum(as.integer(h[17:20]) * 256^(3:0))
      ht <- sum(as.integer(h[21:24]) * 256^(3:0))
      if (ht > 0) w / ht else NA_real_
    } else NA_real_
  }, error = function(e) NA_real_)
  if (!is.na(r) && r > 0) return(r)

  if (requireNamespace("png", quietly = TRUE)) {
    d <- tryCatch(dim(png::readPNG(p)), error = function(e) NULL)
    if (!is.null(d) && length(d) >= 2 && d[1] > 0) return(d[2] / d[1])
  }
  3
}

# pos = "footer" (bas droite des slides de contenu) ou "center" (pages de garde).
# "coin"/"centre" acceptes comme synonymes.
dg_logo <- function(ppt, W, H, pos = "footer", hauteur = NULL, top = NULL, M = 0.3) {
  if (pos == "centre") pos <- "center"
  if (pos == "coin")   pos <- "footer"
  p <- dg_logo_path()
  if (is.null(p)) return(ppt)
  if (is.null(hauteur)) hauteur <- if (pos == "center") 0.75 else 0.32
  larg <- hauteur * dg_logo_ratio(p)
  if (larg > W - 2*M) { larg <- W - 2*M; hauteur <- larg / dg_logo_ratio(p) }
  if (pos == "center") {
    left <- (W - larg) / 2
    if (is.null(top)) top <- 0.35
  } else {
    left <- W - M - larg
    if (is.null(top)) top <- H - 0.12 - hauteur
  }
  ph_with(ppt, value = external_img(p, width = larg, height = hauteur),
          location = ph_location(left = left, top = top, width = larg, height = hauteur))
}

# Liste a puces
dg_puces <- function(txt, size = 11) {
  do.call(block_list, lapply(txt, function(t)
    dg_par(paste0("▪  ", t), size = size, color = DG_TXT)))
}

# Formate un vecteur nomme de taux en "Libelle (xx%)" separes par des ;
dg_liste_champs <- function(v, max_n = 6) {
  if (length(v) == 0) return(NULL)
  v <- sort(v)
  n <- length(v)
  v <- utils::head(v, max_n)
  s <- paste0(names(v), " (", floor(as.numeric(v)), "%)", collapse = " ; ")
  if (n > max_n) s <- paste0(s, " ; ...")
  s
}

# "12,5%" / "12.5%" -> 12.5
dg_num <- function(x) {
  suppressWarnings(as.numeric(gsub(",", ".", gsub("[^0-9,.]", "", as.character(x)))))
}

pm_without_rcsno <- function(df, pm_exclure) {
  if (is.null(df) || nrow(df) == 0 || !("RCSNO" %in% names(df))) return(df[0, , drop = FALSE])

  pm_exclure_clients <- if (!is.null(pm_exclure) && "CLIENT" %in% names(pm_exclure)) {
    unique(trimws(as.character(pm_exclure$CLIENT)))
  } else {
    character(0)
  }
  pm_exclure_clients <- pm_exclure_clients[pm_exclure_clients != "" & !is.na(pm_exclure_clients)]

  liste_pm <- df %>%
    filter(is.na(RCSNO) | trimws(as.character(RCSNO)) == "")

  if ("CLIENT" %in% names(liste_pm) && length(pm_exclure_clients) > 0) {
    liste_pm <- liste_pm %>%
      filter(!(trimws(as.character(CLIENT)) %in% pm_exclure_clients))
  }

  liste_pm %>%
    distinct() %>%
    arrange(AGENCE, EXPL, CLIENT)
}

pp_birthdate_01011900 <- function(df) {
  if (is.null(df) || nrow(df) == 0 || !("DATNAIS" %in% names(df))) return(df[0, , drop = FALSE])

  datnais_chr <- trimws(as.character(df$DATNAIS))
  datnais_chr <- gsub("-", "/", datnais_chr)

  df %>%
    filter(datnais_chr %in% c("01/01/1900", "1900/01/01")) %>%
    distinct() %>%
    arrange(AGENCE, EXPL, CLIENT)
}


    # Repartition notation et taux de complétude des agents par intervalle de taux


    notation=read.csv2(paste0(chemin,"notation.csv"), header = FALSE, row.names = NULL)

    colnames(notation)=c("Agent","Filiale","Note","Date_notation")
    
    filiale_note=unique(notation$Filiale)


notation= notation  %>%
    group_by(Filiale, Note) %>%
    summarise(Count = n(), .groups = "drop") %>%
    group_by(Filiale) %>%
    mutate(Nombre_notes = sum(Count)) %>%
    mutate(Percent = Count / Nombre_notes) %>%    # proportion (0 à  1)
    mutate(Percent = percent(Percent, accuracy = 0.2)) %>%  # transforme en Ã¢â‚¬Å“xx.x%Ã¢â‚¬Â
    select(Filiale, Nombre_notes, Note, Percent) %>%
    pivot_wider(names_from = Note, values_from = Percent, values_fill = "0%")


    


          note_cols <- c("Insuffisant", "Passable", "Bien", "Trà¨s Bien")
          for (col in note_cols) {
            if (!(col %in% colnames(notation))) notation[[col]] <- "0%"
          }






 
for (r in filiale) {
    repertoire=file.path(chemin,r)
    dir.create(repertoire)
}

## Fonctions 

pair <- function(nombre) {
  if (nombre %% 2 == 0) {
    return(TRUE)
  } else {
    return(FALSE)
  }
}

compress <- function(x, sep = " ", dec = ",", digits = 0) {
  formatC(x, format = "f", big.mark = sep, decimal.mark = dec, digits = digits)
}




is_alphanumeric <- function(x) {
  if (is.character(x)) {
    x2 <- suppressWarnings(iconv(x, from = "", to = "UTF-8", sub = ""))
    bad <- is.na(x2)
    if (any(bad)) {
      x2[bad] <- suppressWarnings(iconv(x[bad], from = "latin1", to = "UTF-8", sub = ""))
    }
    x <- x2
  }
  grepl("[a-zA-Z]", x, useBytes = TRUE) | grepl("[0-9]", x, useBytes = TRUE)
}


remove_spaces <- function(x) {
  if (is.character(x)) {
    # Normalise encodage (gà¨re les octets invalides)
    x2 <- suppressWarnings(iconv(x, from = "", to = "UTF-8", sub = ""))
    bad <- is.na(x2)
    if (any(bad)) {
      x2[bad] <- suppressWarnings(iconv(x[bad], from = "latin1", to = "UTF-8", sub = ""))
    }
    return(trimws(gsub("\\t", "", x2, useBytes = TRUE)))
  }
  return(x)
}

# Classe de caracteres "espace" a normaliser : tabulation, espace ordinaire,
# insecable (U+00A0), insecable etroit (U+202F), espace chiffre (U+2007).
# Caracteres litteraux et non echappements \x{...} : PCRE refuse ces derniers
# quand la chaine traitee n'est pas marquee UTF-8 (cas des exports Latin-1).
ESPACES_PARASITES <- paste0("[\t ", intToUtf8(160), intToUtf8(8239), intToUtf8(8199), "]+")

# Colonnes a NE PAS charger, par filiale. Elles sont ecartees des la lecture
# (drop de fread) : elles n'occupent donc jamais de memoire. Une colonne listee
# ici mais absente de l'export est simplement signalee, sans erreur.
COLONNES_IGNOREES <- list(
  MG = c("INTITULE_COMPTE", "LIEU_DELIVRANCE_CIN", "BOITE_POSTALE", "CONSENT_BIC")
)

# Au-dela de ce seuil, read_csv2_big lit le CSV en NB_TRANCHES morceaux
# (mono-thread) au lieu d'une seule passe. Augmenter NB_TRANCHES si la session
# R meurt encore avec une violation d'acces 0xC0000005.
SEUIL_DECOUPAGE_GO <- 1
NB_TRANCHES <- 4

# Nettoyage des espaces (début/fin + multiples)
clean_spaces <- function(x) {
  x <- ifelse(is.na(x), "", as.character(x))
  if (is.character(x)) {
    orig <- x
    x2 <- suppressWarnings(iconv(orig, from = "", to = "UTF-8", sub = ""))
    bad <- is.na(x2)
    if (any(bad)) {
      x2[bad] <- suppressWarnings(iconv(orig[bad], from = "latin1", to = "UTF-8", sub = ""))
    }
    x <- x2
  }
  x <- gsub(ESPACES_PARASITES, " ", x, perl = TRUE)
  trimws(x)
}

# Force AGENCE to 5 chars with leading zeros
pad_agence5 <- function(x) {
  x <- as.character(x)
  x[is.na(x)] <- ""
  x <- trimws(x)
  idx <- x != "" & nchar(x) < 5
  if (any(idx)) {
    x[idx] <- gsub(" ", "0", sprintf("%5s", x[idx]))
  }
  x
}


# Compte le nombre de champs vides sur une liste de colonnes
count_empty_fields <- function(df, fields) {
  if (is.null(df) || length(fields) == 0) return(0L)
  idx <- match(fields, colnames(df))
  idx <- idx[!is.na(idx)]
  if (length(idx) == 0) return(0L)
  sum(sapply(df[, idx, drop = FALSE], function(x) {
    x <- trimws(as.character(x))
    sum(is.na(x) | x == "")
  }))
}

field_completion_rate <- function(df, field, inc = c("")) {
  if (is.null(df) || nrow(df) == 0 || !(field %in% names(df))) return(NA_real_)
  x <- trimws(as.character(df[[field]]))
  rate_floor_no100(sum(is.na(x) | x %in% inc), nrow(df))
}

compute_taux <- function(df, fields) {
  if (is.null(df) || nrow(df) == 0 || length(fields) == 0) return(NA_real_)
  cv <- count_empty_fields(df, fields)
  rate_floor_no100(cv, length(fields) * nrow(df))
}

format_percent <- function(x) {
  if (length(x) == 0 || is.na(x)) "N/A" else paste0(x, "%")
}

min_rate <- function(...) {
  values <- c(...)
  if (length(values) == 0 || all(is.na(values))) NA_real_ else min(values, na.rm = TRUE)
}

# Retire les lignes dont la colonne `col` est vide (ligne gabarit de
# bind_results). df[-which(...), ] renverrait un tableau VIDE si which() ne
# trouve rien (`-integer(0)` selectionne zero ligne) : d'ou ce garde-fou.
drop_lignes_vides <- function(df, col) {
  if (is.null(df) || nrow(df) == 0 || !(col %in% names(df))) return(df)
  idx <- which(trimws(as.character(df[[col]])) == "")
  if (length(idx) == 0) return(df)
  df[-idx, , drop = FALSE]
}

# Table des taux de completude par agent.
# Robuste au cas ou aucun agent n'est concerne (flux vide sur la periode) :
# data.frame() refuse de recycler une valeur de longueur 1 avec des colonnes de
# longueur 0 ("arguments imply differing number of rows: 0, 1").
taux_completude_table <- function(tab, date_periode, flux_stock, pp_pm) {
  n <- if (is.null(tab)) 0L else nrow(tab)
  if (n == 0) {
    return(data.frame(Agents = character(0), Taux = character(0),
                      Date = character(0), flux_stock = character(0),
                      pp_pm = character(0), stringsAsFactors = FALSE))
  }
  # La colonne du taux s'appelle "Taux.de.fiabilisation" apres le mangling de
  # data.frame(), mais peut rester "Taux de fiabilisation" si construite avec
  # check.names = FALSE : on accepte les deux.
  col_taux <- intersect(c("Taux.de.fiabilisation", "Taux de fiabilisation"), names(tab))
  taux <- if (length(col_taux) > 0) tab[[col_taux[1]]] else rep(NA_character_, n)

  data.frame(Agents     = tab$Agents,
             Taux       = taux,
             Date       = rep(as.character(date_periode), n),
             flux_stock = rep(flux_stock, n),
             pp_pm      = rep(pp_pm, n),
             stringsAsFactors = FALSE)
}

get_appreciation <- function(taux, faible, moyen) {
  if (length(taux) == 0 || is.na(taux)) {
    "Aucune donnée"
  } else if (taux < faible) {
    "Faible"
  } else if (taux >= faible & taux < moyen) {
    "Moyen"
  } else {
    "Bon"
  }
}

get_taux_ratt <- function(taux, trimestre, label_retard = TRUE) {
  if (length(taux) == 0 || is.na(taux)) return("N/A")
  taux_rat <- floor(trimestre - taux)
  if (taux_rat < 0) {
    paste("En avance de + ", abs(taux_rat), "%", sep = "")
  } else if (label_retard) {
    paste("En retard de ", taux_rat, "%", sep = "")
  } else {
    paste(taux_rat, "%", sep = "")
  }
}

bind_results <- function(results, template) {
  if (length(results) == 0) return(template[0, , drop = FALSE])
  do.call("rbind", results)
}


reset_row_names <- function(data) {
  if (is.data.frame(data)) row.names(data) <- NULL
  data
}

# Selection des colonnes PP par nom, independamment de leur position dans l'export
select_pp_fields <- function(df, fields) {
  names(df) <- toupper(names(df))
  context_fields <- c("AGENCE", "EXPL", "CLIENT","IDP","DATOUV")
  fields <- toupper(fields)
  selected_fields <- c(context_fields, fields)
  missing_fields <- setdiff(selected_fields, names(df))

  for (field in missing_fields) {
    df[[field]] <- ""
  }

  df <- df[, selected_fields, drop = FALSE]
  names(df)[names(df) == "ORIGINE_REV"] <- "ORIGINE_REVENU"
  df
}

# Selection des colonnes PM par nom, independamment de leur position dans l'export
select_pm_fields <- function(df, fields) {
  names(df) <- toupper(names(df))
  context_fields <- c("AGENCE", "EXPL", "CLIENT","IDM","DATOUV")
  fields <- toupper(fields)
  selected_fields <- c(context_fields, fields)
  missing_fields <- setdiff(selected_fields, names(df))

  for (field in missing_fields) {
    df[[field]] <- ""
  }

  df <- df[, selected_fields, drop = FALSE]
  names(df)[names(df) == "ORIGINE_REV"] <- "ORIGINE_REVENU"
  df
}

# Nettoie le fichier source si "RCS NÃ‚Â°" empàªche la lecture
sanitize_rcsno_file <- function(path) {
  if (!file.exists(path)) return(path)
  txt <- readLines(path, encoding = "Latin1", warn = FALSE, skipNul = TRUE)
  Encoding(txt) <- "latin1"
  pat1 <- paste0("RCS N", "\xB0")       # degree in Latin1
  pat2 <- paste0("RCS N", "\xC2\xB0")   # UTF-8 degree read as Latin1 (??)
  if (!any(grepl(pat1, txt, fixed = TRUE, useBytes = TRUE) | grepl(pat2, txt, fixed = TRUE, useBytes = TRUE))) return(path)
  tmp <- paste0(path, ".clean")
  txt <- gsub(pat1, "RCS N", txt, fixed = TRUE, useBytes = TRUE)
  txt <- gsub(pat2, "RCS N", txt, fixed = TRUE, useBytes = TRUE)
  writeLines(txt, tmp, useBytes = TRUE)
  tmp
}

# Parsing robuste des dates (YYYY-MM-DD ou DD/MM/YYYY)
parse_date_any <- function(x) {
  x <- trimws(as.character(x))
  x[x == ""] <- NA
  out <- suppressWarnings(ymd(x))
  idx <- is.na(out)
  if (any(idx)) out[idx] <- suppressWarnings(dmy(x[idx]))
  out
}

# Applique un style par lot selon une colonne de statut
apply_status_style <- function(wb, sheet, data, value_col,
                               col_good = value_col, col_mid = value_col, col_low = value_col,
                               good = "Bon", mid = "Moyen", header_rows = 1,
                               style_good = bonStyle, style_mid = moyenStyle, style_low = lowStyle) {
  if (is.null(data) || nrow(data) == 0) return(invisible(NULL))
  v <- data[[value_col]]
  rows_good <- which(v == good) + header_rows
  if (length(rows_good) > 0) {
    addStyle(wb, sheet, style_good, cols = col_good, rows = rows_good, gridExpand = FALSE, stack = TRUE)
  }
  rows_mid <- which(v == mid) + header_rows
  if (length(rows_mid) > 0) {
    addStyle(wb, sheet, style_mid, cols = col_mid, rows = rows_mid, gridExpand = FALSE, stack = TRUE)
  }
  rows_low <- which(!(v %in% c(good, mid)) & !is.na(v)) + header_rows
  if (length(rows_low) > 0) {
    addStyle(wb, sheet, style_low, cols = col_low, rows = rows_low, gridExpand = FALSE, stack = TRUE)
  }
  invisible(NULL)
}






# Fonction pour nettoyer les noms de colonnes (BOM et mauvais encodage)
clean_colnames <- function(data) {
    # Supprime les caractà¨res de BOM et mauvais encodage au début
    names(data) <- gsub("^[\\xef\\xbb\\xbf\\u00ef\\u00bb\\u00bf]|^à¯\\.\\.\\.?|^à¯\\.\\.|^à¯\\.\\xef", "", names(data))
    return(data)
}

# Detection automatique d'encodage (heuristique rapide)
detect_encoding <- function(path, n = 100000) {
  con <- file(path, "rb")
  on.exit(close(con), add = TRUE)
  raw <- readBin(con, "raw", n = n)

  if (length(raw) >= 3 &&
      raw[1] == as.raw(0xEF) &&
      raw[2] == as.raw(0xBB) &&
      raw[3] == as.raw(0xBF)) {
    return("UTF-8-BOM")
  }

  raw_nz <- raw[raw != as.raw(0x00)]
  if (length(raw_nz) == 0) {
    return("UTF-8")
  }

  txt <- tryCatch(
    iconv(rawToChar(raw_nz), from = "UTF-8", to = "UTF-8", sub = NA),
    error = function(e) NA
  )
  if (!is.na(txt)) {
    return("UTF-8")
  }

  return("Windows-1252")
}

# Lecture CSV2 avec encodage auto + correction si besoin


# Nettoyage post-lecture : enlà¨ve les guillemets, préserve accents/trémas/@
clean_quotes_df <- function(df) {
  char_cols <- vapply(df, is.character, logical(1))
  df[char_cols] <- lapply(df[char_cols], function(x) {
    # Sécurise l'encodage (préserve é, à¯, @, etc.)
    x <- enc2utf8(x)
    # Supprime tous les guillemets doubles
    x <- gsub('"', "", x, fixed = TRUE)
    # Nettoie les espaces résiduels en début/fin
    trimws(x)
  })
  # Nettoie aussi les noms de colonnes
  names(df) <- trimws(gsub('"', "", enc2utf8(names(df)), fixed = TRUE))
  df
}

# Lecture CSV2 avec encodage auto + correction si besoin
read_csv2_auto <- function(path, ...) {
  enc_guess <- detect_encoding(path)
  encodings <- unique(c(enc_guess, "UTF-8-BOM", "UTF-8", "Windows-1252", "Latin1"))
  last_error <- NULL
  dots <- list(...)
  if (!"strip.white" %in% names(dots)) dots$strip.white <- TRUE
  # Force la lecture en caractà¨res (évite les conversions en facteurs)
  if (!"stringsAsFactors" %in% names(dots)) dots$stringsAsFactors <- FALSE

  for (enc in encodings) {
    res <- tryCatch(
      withCallingHandlers(
        do.call(read.csv2, c(list(path, fileEncoding = enc), dots)),
        warning = function(w) {
          msg <- conditionMessage(w)
          if (grepl("invalid input|input string|entrée incorrecte|entree incorrecte", msg, ignore.case = TRUE)) {
            invokeRestart("muffleWarning")
            stop("invalid_input_warning")
          }
          if (grepl("incomplete final line", msg, ignore.case = TRUE)) {
            invokeRestart("muffleWarning")
          }
          if (grepl("EOF", msg, ignore.case = TRUE) &&
              grepl("quoted", msg, ignore.case = TRUE)) {
            invokeRestart("muffleWarning")
            stop("eof_quoted_string")
          }
        }
      ),
      error = function(e) e
    )

    if (!inherits(res, "error")) {
      if (!identical(enc, enc_guess)) {
        message("Correction encodage: ", enc)
      }
      return(clean_quotes_df(res))
    }

    last_error <- res

    # Retry sans guillemets si EOF dans chaà®ne quotée ou entrée invalide
    msg <- conditionMessage(res)
    if (grepl("eof_quoted_string|invalid_input_warning|invalid input|input string|entrée incorrecte|entree incorrecte", msg, ignore.case = TRUE)) {
      dots2 <- dots
      if (!"quote" %in% names(dots2)) dots2$quote <- ""
      if (!"fill" %in% names(dots2)) dots2$fill <- TRUE
      res2 <- tryCatch(
        do.call(read.csv2, c(list(path, fileEncoding = enc), dots2)),
        error = function(e) e
      )
      if (!inherits(res2, "error")) {
        if (!identical(enc, enc_guess)) {
          message("Correction encodage: ", enc)
        }
        message("Correction lecture: guillemets desactives (quote='')")
        return(clean_quotes_df(res2))
      }
    }
  }

  stop("Lecture impossible (encodage). Dernier avertissement: ", conditionMessage(last_error))
}

# Lecture rapide des gros fichiers (pp_stock/pm_stock ~2 Go) via data.table::fread
# - multi-thread, strip.white natif (espaces parasites dans les colonnes)
# - cache .rds : les relances du script rechargent en quelques secondes
# - type.convert reproduit le typage de read.csv2 (dec = ",")
read_csv2_big <- function(path, cache = TRUE, drop = NULL) {
  t0 <- Sys.time()
  # Le cache depend des colonnes ignorees : suffixe distinct pour ne pas
  # recharger un cache complet quand on demande une lecture allegee (ou l'inverse).
  # Suffixe ".v2q" : les exports filiales entourent desormais les champs texte
  # de guillemets. On change la cle de cache pour ne PAS recharger un ancien
  # .rds construit sur la version non quotee (ADRESSE tronquee sur les ";").
  rds <- paste0(path, if (length(drop)) "-drop" else "", ".v2q.rds")
  if (cache && file.exists(rds) && file.mtime(rds) >= file.mtime(path)) {
    message("[read_csv2_big] Lecture depuis le cache: ", basename(rds))
    df <- readRDS(rds)
    message("[read_csv2_big] Cache charge en ",
            round(difftime(Sys.time(), t0, units = "secs")), "s (",
            nrow(df), " lignes)")
    return(df)
  }

  message("[read_csv2_big] ", basename(path), " : ",
          round(file.size(path) / 1024^2), " Mo, ",
          data.table::getDTthreads(), " threads fread")

  enc_guess <- detect_encoding(path)
  enc_fread <- if (grepl("UTF-?8", enc_guess, ignore.case = TRUE)) "UTF-8" else "Latin-1"
  message("[read_csv2_big] Encodage fread: ", enc_fread)

  # fill = Inf : certaines lignes ont des champs en trop (";" dans un libellé),
  # sans cela fread s'arràªte en cours de fichier (lecture TRONQUEE silencieuse).
  # Le warning "Stopped early" est donc traité comme un échec.
  fread_try <- function(q, skip = 0L, nrows = Inf, col_names = NULL, nthread = NULL,
                        dropcols = NULL) {
    stopped <- FALSE
    args <- list(path, sep = ";", dec = ",", quote = q,
                 colClasses = "character", encoding = enc_fread,
                 strip.white = TRUE, fill = Inf, blank.lines.skip = TRUE,
                 showProgress = TRUE, data.table = TRUE, nrows = nrows)
    if (!is.null(nthread)) args$nThread <- nthread
    # drop par INDICE et non par nom : en mode tranche les colonnes sortent en
    # V1, V2, ... (header = FALSE), un drop par nom ne matcherait rien.
    if (length(dropcols)) args$drop <- dropcols
    if (!is.null(col_names)) {
      # Tranche : on saute l'en-tete + les lignes deja lues (sinon la 1re ligne
      # de donnees servirait d'en-tete et la tranche perdrait une ligne).
      # NE PAS passer col.names ici : cela figerait le nombre de colonnes et
      # les lignes ayant des champs en trop (";" dans un libelle) deborderaient
      # sur une NOUVELLE LIGNE au lieu de creer des colonnes V4, V5...
      # Les colonnes sortent donc en V1, V2, ... et sont renommees apres le
      # recollage, ce qui reproduit exactement la lecture en une passe.
      args$skip <- skip + 1L
      args$header <- FALSE
    }
    df <- withCallingHandlers(
      tryCatch(
        do.call(data.table::fread, args),
        error = function(e) {
          message("[read_csv2_big] fread (quote=", deparse(q), ") a echoue: ",
                  conditionMessage(e))
          NULL
        }
      ),
      warning = function(w) {
        msg <- conditionMessage(w)
        if (grepl("Stopped early|Discarded single-line", msg)) {
          message("[read_csv2_big] Lecture incomplete (quote=", deparse(q), "): ", msg)
          stopped <<- TRUE
        }
        invokeRestart("muffleWarning")
      }
    )
    if (stopped) return(NULL)
    df
  }

  # Répare les octets invalides (fichiers UTF-8 contenant des restes Latin-1),
  # enlà¨ve les guillemets et espaces résiduels, puis retype comme read.csv2.
  # validEnc (test C rapide) évite de convertir chaque chaà®ne : seules les
  # valeurs réellement invalides passent par iconv.
  fix_utf8 <- function(x) {
    bad <- !validEnc(x)
    if (any(bad)) {
      x[bad] <- suppressWarnings(iconv(x[bad], from = "latin1", to = "UTF-8", sub = ""))
    }
    x
  }

  # Nettoyage colonne par colonne PAR REFERENCE (data.table::set).
  # Un "df[] <- lapply(df, ...)" dupliquerait tout le tableau en memoire : sur
  # 3 Go de CSV (soit >10 Go une fois en RAM) cela suffit a faire crasher R
  # avec une violation d'acces 0xC0000005. Ici une seule colonne est en double
  # a un instant donne, et l'ancienne est liberee immediatement.
  # retype = FALSE pour les tranches : le type.convert final se fait une seule
  # fois sur le tableau recolle (sinon deux tranches peuvent recevoir des types
  # differents selon leur contenu, et le rbind echoue ou coerce en silence).
  clean_dt <- function(df, retype = TRUE) {
    for (j in seq_len(ncol(df))) {
      x <- df[[j]]
      x <- fix_utf8(x)
      if (any(grepl('"', x, fixed = TRUE))) {
        x <- gsub('"', "", x, fixed = TRUE)
      }
      # Espaces parasites de l'export : tabulations, espaces insecables et
      # suites d'espaces internes -> un seul espace, puis trim.
      x <- gsub(ESPACES_PARASITES, " ", x, perl = TRUE)
      x <- trimws(x)
      if (retype) x <- type.convert(x, as.is = TRUE, dec = ",")
      data.table::set(df, j = j, value = x)
      rm(x)
      if (j %% 10 == 0) gc(verbose = FALSE)
    }
    gc(verbose = FALSE)
    data.table::setnames(df, trimws(gsub('"', "", fix_utf8(names(df)), fixed = TRUE)))
    df
  }

  taille_go <- file.size(path) / 1024^3

  # ---- Colonnes a ignorer : resolution des noms en INDICES ----
  # Faite sur l'en-tete reel du fichier : une colonne absente de cet export est
  # simplement signalee et ignoree (pas d'erreur), et le drop est ensuite
  # utilisable aussi bien en lecture directe qu'en mode tranche.
  drop_idx <- NULL
  entete_all <- NULL
  if (length(drop)) {
    entete_all <- fread_try("\"", nrows = 0L)
    if (is.null(entete_all) || ncol(entete_all) < 2) entete_all <- fread_try("", nrows = 0L)
    if (!is.null(entete_all)) {
      noms_reels <- trimws(gsub('"', "", names(entete_all), fixed = TRUE))
      drop_idx <- which(toupper(noms_reels) %in% toupper(drop))
      absentes <- setdiff(toupper(drop), toupper(noms_reels))
      if (length(absentes)) {
        message("[read_csv2_big] Colonnes a ignorer absentes du fichier: ",
                paste(absentes, collapse = ", "))
      }
      if (length(drop_idx)) {
        message("[read_csv2_big] Colonnes ignorees (non chargees): ",
                paste(noms_reels[drop_idx], collapse = ", "))
      } else {
        drop_idx <- NULL
      }
    }
  }

  if (taille_go <= SEUIL_DECOUPAGE_GO) {
    # ---- Lecture en une seule passe (fichiers "normaux") ----
    # quote="\"" en premier : les exports filiales entourent desormais les
    # champs texte de guillemets, ce qui protege les ";" internes (ADRESSE).
    # Repli quote="" si les guillemets sont mal apparies (fichier lu en 1 colonne).
    df <- fread_try("\"", dropcols = drop_idx)
    if (is.null(df) || ncol(df) < 2) {
      message("[read_csv2_big] Nouvel essai sans gestion des guillemets ...")
      df <- fread_try("", dropcols = drop_idx)
      if (!is.null(df) && ncol(df) < 2) df <- NULL
    }
    if (is.null(df)) {
      message("[read_csv2_big] ATTENTION: bascule sur read_csv2_auto (LENT, plusieurs minutes)")
      return(read_csv2_auto(path))
    }
    message("[read_csv2_big] fread OK: ", nrow(df), " lignes en ",
            round(difftime(Sys.time(), t0, units = "secs")), "s. Typage...")
    df <- clean_dt(df, retype = TRUE)

  } else {
    # ---- Lecture par tranches (fichiers > SEUIL_DECOUPAGE_GO) ----
    # Chaque tranche est lue puis nettoyee avant de passer a la suivante : le
    # pic memoire du nettoyage porte sur 1/N du fichier au lieu du tout.
    # nThread = 1 : le parsing multi-thread de fread sur fichier malforme
    # (champs en trop, guillemets mal apparies) est une cause connue de crash
    # 0xC0000005 ; la lecture est plus lente mais ne tue pas la session.
    message("[read_csv2_big] Fichier de ", round(taille_go, 2),
            " Go > ", SEUIL_DECOUPAGE_GO, " Go : lecture en ",
            NB_TRANCHES, " tranches, mono-thread.")

    # En-tete APRES drop : col_names doit decrire les colonnes reellement lues,
    # puisque les tranches sortent en V1, V2, ... et sont renommees dessus.
    entete <- fread_try("\"", nrows = 0L, dropcols = drop_idx)
    quote_ok <- "\""
    if (is.null(entete) || ncol(entete) < 2) {
      entete <- fread_try("", nrows = 0L, dropcols = drop_idx)
      quote_ok <- ""
    }
    if (is.null(entete) || ncol(entete) < 2) {
      message("[read_csv2_big] ATTENTION: en-tete illisible, bascule sur read_csv2_auto")
      return(read_csv2_auto(path))
    }
    col_names <- names(entete)

    # Taille de tranche ESTIMEE a partir de la longueur moyenne des 1000
    # premieres lignes. Volontairement pas de comptage exact prealable : un
    # fread de comptage sur fichier malforme s'arrete tot et sous-compte, ce
    # qui tronquerait les donnees en silence. La boucle ci-dessous lit donc
    # jusqu'a epuisement reel du fichier ; l'estimation ne joue que sur le
    # nombre de tranches, jamais sur l'exhaustivite.
    ech <- readLines(path, n = 1000L, warn = FALSE)
    oct_par_ligne <- max(1, mean(nchar(ech, type = "bytes") + 1))
    n_estime <- max(1, ceiling(file.size(path) / oct_par_ligne))
    par_tranche <- max(1L, as.integer(ceiling(n_estime / NB_TRANCHES)))
    message("[read_csv2_big] ~", n_estime, " lignes estimees -> ",
            par_tranche, " par tranche")

    morceaux <- list()
    i <- 0L
    repeat {
      i <- i + 1L
      saut <- (i - 1L) * par_tranche
      message("[read_csv2_big] Tranche ", i, " (a partir de la ligne ",
              saut + 1L, ") ...")
      tr <- fread_try(quote_ok, skip = saut, nrows = par_tranche,
                      col_names = col_names, nthread = 1L, dropcols = drop_idx)
      if (is.null(tr)) {
        message("[read_csv2_big] Tranche ", i, " illisible, bascule sur read_csv2_auto")
        return(read_csv2_auto(path))
      }
      if (nrow(tr) == 0L) break          # fin de fichier atteinte
      morceaux[[length(morceaux) + 1L]] <- clean_dt(tr, retype = FALSE)
      lu <- nrow(tr)
      rm(tr); gc(verbose = FALSE)
      if (lu < par_tranche) break        # derniere tranche (incomplete)
    }

    message("[read_csv2_big] Recollage des ", length(morceaux), " tranches ...")
    # use.names = TRUE sur des noms positionnels (V1, V2, ...) : les tranches
    # n'ayant pas de ligne malformee ont moins de colonnes, fill = TRUE complete
    # alors par NA, comme le ferait la lecture en une passe.
    df <- data.table::rbindlist(morceaux, use.names = TRUE, fill = TRUE)
    rm(morceaux); gc(verbose = FALSE)

    # Restitution des vrais noms de colonnes ; les colonnes surnumeraires
    # (champs en trop) gardent leur nom positionnel V4, V5, ... comme en une passe.
    n_nommees <- min(length(col_names), ncol(df))
    if (n_nommees > 0L) {
      data.table::setnames(df, seq_len(n_nommees), col_names[seq_len(n_nommees)])
    }

    # Les NA presents ici ne peuvent venir que du fill de rbindlist (colonnes
    # surnumeraires absentes de certaines tranches) : le typage n'a pas encore
    # eu lieu, donc un champ vide du CSV vaut "" et non NA. On restitue "",
    # ce que produit la lecture en une passe.
    for (j in seq_len(ncol(df))) {
      x <- df[[j]]
      if (is.character(x) && anyNA(x)) {
        x[is.na(x)] <- ""
        data.table::set(df, j = j, value = x)
      }
    }

    # Typage final, une seule fois, sur le tableau complet
    for (j in seq_len(ncol(df))) {
      data.table::set(df, j = j,
                      value = type.convert(df[[j]], as.is = TRUE, dec = ","))
      if (j %% 10 == 0) gc(verbose = FALSE)
    }
    gc(verbose = FALSE)
    message("[read_csv2_big] Recollage OK: ", nrow(df), " lignes")
  }

  data.table::setDF(df)   # par reference : pas de copie

  if (cache) {
    message("[read_csv2_big] Ecriture du cache ", basename(rds), " ...")
    # compress = FALSE : gzip mono-thread prendrait plusieurs minutes sur 2 Go
    tryCatch(saveRDS(df, rds, compress = FALSE), error = function(e)
      message("[read_csv2_big] Cache non ecrit (", conditionMessage(e), ")"))
  }
  message("[read_csv2_big] Termine en ",
          round(difftime(Sys.time(), t0, units = "secs")), "s")
  df
}


# Agrege les lignes partageant le meme CLIENT en une seule ligne.
# Pour chaque colonne : valeurs distinctes non vides, separees par une virgule.
#   CLIENT 42 | TEL "0102"  | VILLE "DAKAR"   ->  CLIENT 42 | TEL "0102,0304"
#   CLIENT 42 | TEL "0304"  | VILLE "DAKAR"       VILLE "DAKAR"
# Les colonnes constantes a l'interieur de chaque groupe gardent leur type
# d'origine (numerique, date...) ; seules celles reellement fusionnees passent
# en texte, sinon un "1,2" issu de deux valeurs 1 et 2 serait relu comme le
# nombre 1,2 (dec = ",") et corromprait la donnee.
agreger_par_client <- function(df, cle = "CLIENT") {
  if (is.null(df) || !(cle %in% names(df)) || nrow(df) == 0) return(df)

  ordre <- names(df)
  data.table::setDT(df)

  n_doublons <- sum(duplicated(df[[cle]]))
  if (n_doublons == 0L) {
    message("[agreger_par_client] Aucun ", cle, " en double : rien a agreger.")
    data.table::setDF(df)
    return(df)
  }
  message("[agreger_par_client] ", n_doublons, " ligne(s) en double sur ", cle,
          " -> agregation ...")

  cols <- setdiff(names(df), cle)

  # 1 seule passe groupee pour savoir quelles colonnes varient dans un groupe
  nu <- df[, lapply(.SD, data.table::uniqueN), by = c(cle), .SDcols = cols]
  varie <- vapply(cols, function(cn) any(nu[[cn]] > 1L), logical(1))
  rm(nu); gc(verbose = FALSE)

  cols_fusion <- cols[varie]
  cols_stables <- cols[!varie]
  message("[agreger_par_client] ", length(cols_fusion),
          " colonne(s) fusionnee(s) en texte",
          if (length(cols_fusion)) paste0(" : ", paste(cols_fusion, collapse = ", ")) else "")

  fusionner <- function(x) {
    v <- trimws(as.character(x))
    v <- unique(v[!is.na(v) & nzchar(v)])
    if (length(v) == 0L) "" else paste(v, collapse = ",")
  }

  # Colonnes stables : on garde la 1re valeur (donc le type d'origine)
  res <- df[, lapply(.SD, data.table::first), by = c(cle), .SDcols = cols_stables]
  if (length(cols_fusion)) {
    res_f <- df[, lapply(.SD, fusionner), by = c(cle), .SDcols = cols_fusion]
    res <- merge(res, res_f, by = cle, sort = FALSE)
    rm(res_f); gc(verbose = FALSE)
  }

  data.table::setcolorder(res, intersect(ordre, names(res)))
  data.table::setDF(res)
  message("[agreger_par_client] ", nrow(df), " lignes -> ", nrow(res),
          " (", cle, " uniques)")
  res
}

# Regroupe les lignes partageant le meme CLIENT en une seule ligne.
# Pour chaque colonne : valeurs distinctes non vides, separees par une virgule.
#   - une seule valeur distincte  -> la valeur telle quelle (type d'origine conserve)
#   - plusieurs valeurs           -> "VAL1,VAL2" (la colonne devient character)
#   - aucune valeur non vide      -> "" (ou NA si la colonne n'est pas character)
# L'ordre d'apparition des CLIENT est conserve (by=, et non keyby= qui trierait).
#   - colonnes de date_cols     -> la date la PLUS ANCIENNE (pas de concatenation),
#                                  reformatee en JJ/MM/AAAA pour rester lisible
#                                  par le dmy() applique plus loin dans le script.
agreger_doublons <- function(df, cle = "CLIENT", date_cols = "DATOUV") {
  if (is.null(df) || nrow(df) == 0 || !(cle %in% names(df))) return(df)

  n_avant <- nrow(df)
  n_cles  <- data.table::uniqueN(df[[cle]])
  if (n_cles == n_avant) {
    message("[agreger_doublons] Aucun ", cle, " en double (", n_avant, " lignes)")
    return(df)
  }

  message("[agreger_doublons] ", n_avant - n_cles, " ligne(s) en trop sur ",
          n_avant, " : regroupement sur ", n_cles, " ", cle, " uniques ...")

  ordre_cols <- names(df)
  dt <- data.table::as.data.table(df)

  # Fusion TOUJOURS en character : data.table impose un type homogene par
  # colonne entre les groupes, or un groupe sans doublon rendrait le type
  # d'origine et un groupe avec doublons du texte -> erreur de type.
  fusion <- function(x) {
    x <- as.character(x)
    u <- unique(x[!is.na(x) & trimws(x) != ""])
    if (length(u) == 0L) return("")
    paste(u, collapse = ",")
  }

  res <- dt[, lapply(.SD, fusion), by = c(cle)]

  # Dates : on ne concatene pas, on garde la PLUS ANCIENNE ouverture du client.
  # Recalcul sur dt (valeurs d'origine) et non sur res, dont la colonne contient
  # deja les dates concatenees par fusion().
  for (dc in intersect(date_cols, names(res))) {
    agg <- dt[, .(..min_d = {
      d <- parse_date_any(get(dc))
      if (all(is.na(d))) NA_character_ else format(min(d, na.rm = TRUE), "%d/%m/%Y")
    }), by = c(cle)]
    # Jointure sur la cle : ne depend pas de l'ordre des lignes.
    idx <- match(res[[cle]], agg[[cle]])
    val <- agg$..min_d[idx]
    val[is.na(val)] <- ""
    data.table::set(res, j = dc, value = val)
  }

  # Retour au type d'origine pour les colonnes qui n'ont finalement subi aucune
  # concatenation : sans cela, des colonnes numeriques (CA, CAPITAL...)
  # deviendraient du texte pour toute la filiale.
  for (nm in setdiff(names(res), cle)) {
    if (!is.character(df[[nm]]) && !any(grepl(",", res[[nm]], fixed = TRUE))) {
      data.table::set(res, j = nm,
                      value = type.convert(res[[nm]], as.is = TRUE))
    }
  }

  data.table::setcolorder(res, intersect(ordre_cols, names(res)))
  data.table::setDF(res)

  message("[agreger_doublons] Termine : ", nrow(res), " lignes")
  res
}

# Excel sheet names must be <= 31 chars and avoid special chars.
safe_sheet_name <- function(name, max_len = 31) {
  if (is.na(name) || !nzchar(name)) name <- "Sheet"
  name <- gsub("[:\\\\/\\?\\*\\[\\]]", " ", name)
  name <- gsub("[\r\n\t]+", " ", name)
  name <- trimws(name)
  if (nchar(name) > max_len) name <- substr(name, 1, max_len)
  if (!nzchar(name)) name <- "Sheet"
  name
}

unique_sheet_name <- function(wb, name) {
  base <- safe_sheet_name(name)
  existing <- tryCatch(openxlsx::getSheetNames(wb), error = function(e) NULL)
  if (is.null(existing) && !is.null(wb$worksheets)) {
    existing <- names(wb$worksheets)
  }
  if (is.null(existing)) return(base)
  candidate <- base
  i <- 1
  while (candidate %in% existing) {
    suffix <- paste0("_", i)
    candidate <- base
    if (nchar(candidate) + nchar(suffix) > 31) {
      candidate <- substr(candidate, 1, 31 - nchar(suffix))
    }
    candidate <- paste0(candidate, suffix)
    i <- i + 1
  }
  candidate
}



#Gestion des couleurs

vert_fonce="#22963B"
vert_clair="#E0F7DC"  
orange="#FFBD0E"
rouge="#F63C00"
bleu="#C2E7FB"

headerStyle <- createStyle(
  fontSize = 14, fontColour = "white", halign = "left",
  fgFill = "#09982E", border = "TopBottom", borderColour = "black"
  )
bonStyle <- createStyle(halign = "left",
fgFill = "#298904"
)
moyenStyle <- createStyle(halign = "left",
fgFill = "#F8DF19"
)
lowStyle <- createStyle(halign = "left",
fgFill = "#D90A0A"
)


## Initiation des tableaux

Sys.setlocale("LC_TIME", "fr_FR.UTF-8")

      cat("ETAPE 1/3: TAUX DE FIABILISATION PAR AGENT \n")


tableau_fiabilisation=data.frame(Flux=c("","Taux de fiabilisation","Appréciation"))
tableau_fiabilisation_stock=data.frame(Stock=c("","Taux de fiabilisation","Appréciation"))

minimum_pp=c()
minimum_pm=c()

fiabilisation_fil=c()
ppt_dg=read_pptx(paste(chemin,"Rapport de suivi.pptx",sep=""))



production = function(fil) {

# Décrompession des fichiers
     sigle=paste("BOA_",fil, sep="")

  
            cat("######################################################\n")
            cat("######### Extraction des données de \n",sigle, "#######\n")
            cat("#####################################################\n")

	chemin_fichier_zip <- paste0("C:\\KYC_Filiales\\sftpkycboa",tolower(fil),"\\INKYC\\",fil,".7z")
	dossier_sortie <- paste0("C:\\Fiabilisation KYC\\R\\",fil,"\\data")
	mot_de_passe <- id[2,2]
	# On extrait tout sauf le fichier de scoring (conservé tel quel dans //data//)
	fichier_exclu <- paste0("scoring_", fil, ".csv")
	commande <- sprintf('%s x "%s" -o"%s" -p"%s" -xr!"%s" -y',chemin_7z, chemin_fichier_zip, dossier_sortie, mot_de_passe, fichier_exclu)
	system(commande)
  # Traitements des données filiales initiales

   ## Clients non fiabilisables

    non_fiab=read.csv2(paste(sep="",chemin,fil,"//data//non_fiab.csv"), fileEncoding = "UTF-8-BOM")
    non_fiab <- clean_colnames(non_fiab)
    non_fiab=as.data.frame(lapply(non_fiab, remove_spaces))

    non_fiab=non_fiab[non_fiab$CODE=="OI",]


# ============================
# 1. Chemin du fichier
# ============================
chemin_pm <- file.path(chemin, fil, "data", "pm_stock.csv")

# Vérification que le fichier existe avant lecture
if (!file.exists(chemin_pm)) {
  stop("Fichier introuvable : ", chemin_pm)
}

# ============================
# 2. Fonction de détection d'encodage (utilisée par read_csv2_auto)
# ============================
detect_encoding <- function(path, n = 100000) {
  raw_sample <- readBin(path, what = "raw", n = n)
  encodings <- stri_enc_detect(raw_sample)
  encoding_detected <- encodings[[1]]$Encoding[1]

  # Sécurité si la détection échoue
  if (is.null(encoding_detected) || is.na(encoding_detected)) {
    encoding_detected <- "UTF-8"
  }

  cat("Encodage retenu :", encoding_detected, "\n")
  encoding_detected
}

# ============================
# 3. Import du CSV
# ============================
# read_csv2_big : lecture rapide fread (multi-thread) + cache .rds,
# gère encodage, espaces parasites, guillemets ; fallback read_csv2_auto
cols_ignorees <- COLONNES_IGNOREES[[fil]]   # NULL si la filiale n'en declare pas

pm_stock <- read_csv2_big(chemin_pm, drop = cols_ignorees)

# Un meme CLIENT peut apparaitre sur plusieurs lignes : on les regroupe des le
# depart, en concatenant par une virgule les valeurs differentes de chaque colonne.
pm_stock <- agreger_doublons(pm_stock, "CLIENT")
pm_stock=unique(pm_stock )

# ============================
# 3 bis. Clients PM exclus du controle RCSNO
# ============================
# Codes AGEC lus dans la colonne pm_exclure de prerequis.csv. La liste des
# CLIENT concernes ne depend que de pm_stock : on la calcule une seule fois ici,
# a l'import, au lieu de la reconstruire au moment du rapport Direction generale.
pm_exclure_codes <- pm_agec_exclus(fil)
cat("######### AGEC exclus du controle RCSNO pour", fil, ":",
    paste(pm_exclure_codes, collapse = ", "), "\n")

pm_exclure <- (function(df, codes) {
  df <- clean_colnames(df)
  names(df) <- toupper(names(df))
  if (!all(c("AGEC", "CLIENT") %in% names(df))) return(df[0, , drop = FALSE])
  df[trimws(toupper(as.character(df$AGEC))) %in% codes, , drop = FALSE]
})(pm_stock, pm_exclure_codes)
cat("#########", nrow(pm_exclure), "client(s) PM exclus du controle RCSNO\n")

# La periode doit etre connue avant le filtrage du flux PM (plus bas), donc
# avant le chargement complet de pp_stock. On lit ici la seule colonne DATOUV
# du fichier PP : lecture rapide, sans toucher au cache de read_csv2_big.
datouv_pp <- tryCatch({
  f_pp <- file.path(chemin, fil, "data", "pp_stock.csv")
  if (file.exists(f_pp)) {
    data.table::fread(f_pp, sep = ";", select = "DATOUV",
                      showProgress = FALSE, nThread = 1)$DATOUV
  } else character(0)
}, error = function(e) {
  cat("######### ATTENTION", fil,
      ": DATOUV de pp_stock.csv illisible, periode calee sur les PM seules\n")
  character(0)
})

dates_dispo <- sort(unique(c(dmy(pm_stock$DATOUV), dmy(datouv_pp))), decreasing = TRUE)
dates_triees <- dates_dispo   # conserve pour compatibilite

# Bornes de la periode consideree pour cette filiale (cf. bornes_periode /
# colonne PERIODE de prerequis.csv).
# premier_jour = premier jour de la periode (inclus)
# jour_recent  = lendemain du dernier jour de la periode (exclu)
periode <- bornes_periode(fil, dates_dispo)
premier_jour  <- periode$debut
jour_recent   <- periode$fin
periode_libelle <- periode$libelle
periode_mode_fil <- periode$mode


## Traitement des PM

if (fil=="ML") {

pm_stock_s=pm_stock

pm_stock_s$CLIENT=as.numeric(pm_stock_s$CLIENT)

risque_s=read.csv2(paste(sep="",chemin,"ML//data//risque.csv"),header=F)
risque_s[,1]=as.numeric(risque_s[,1])
risque=data.frame(CLIENT=risque_s[,1], RISQUE=risque_s[,3])

pm_stock_s=left_join(pm_stock_s,risque, by='CLIENT')



pm_stock_s$RISQUE.x = pm_stock_s$RISQUE.y


pm_stock_s=pm_stock_s[,-which(colnames(pm_stock_s)=="RISQUE.y")]

colnames(pm_stock_s)[which(colnames(pm_stock_s)=="RISQUE.x")] = "RISQUE"

pm_stock_s$RESID[pm_stock_s$PAYS_JUR=="MALI"]="O"
pm_stock_s$RESID[pm_stock_s$PAYS_JUR!="MALI"]="N"

pm_stock = pm_stock_s


} else {
   pm_stock=pm_stock
}




# Vérification
nrow(pm_stock)
colnames(pm_stock)


if ("AGENCE_LIB" %in% names(pm_stock)) {
  colnames(pm_stock)[which(names(pm_stock)=="AGENCE_LIB")] <- "AGENCELIB"

} 

if ("LIB_AGENCE" %in% names(pm_stock)) {
   colnames(pm_stock)[which(names(pm_stock)=="LIB_AGENCE")] <- "AGENCELIB"
}

   
                    
    pm_stock[is.na(pm_stock)]=""

    pm_stock=as.data.frame(lapply(pm_stock, remove_spaces))

    if ("AGENCE" %in% names(pm_stock)) {
        pm_stock$AGENCE <- pad_agence5(pm_stock$AGENCE)
    }

    pm_stock$DATOUV = dmy(pm_stock$DATOUV) 
    
    diff_s=setdiff(pm_stock$CLIENT, non_fiab$CLIENT)

    
    diff_sec_pm=intersect(non_fiab$CLIENT, pm_stock$CLIENT)

    pm_stock= pm_stock[pm_stock$CLIENT %in% diff_s,]

    # Devise(s) de la filiale : colonne "Devise" de prerequis.csv, adressee par
    # nom (elle etait auparavant lue par indice, ce qui cassait au moindre
    # ajout de colonne). Plusieurs devises possibles, separees par une virgule.
    devise_fil = strsplit(prerequis_valeur(fil, "devise"), split = ",")

    clients_devise_pm = which(pm_stock$DEVISE %in% devise_fil)

    pm_stock[clients_devise_pm,"DEVISE"]= ""

 

    pm_flux <- pm_stock %>%
    filter(
        DATOUV >= as.Date(premier_jour),
        DATOUV < as.Date(jour_recent)
    )

  

    cols_vides_pm <- c("AGEC","CODAPE","RCSNO","CAPITAL","CA","RESULTAT","ORIGINE_REVENU","TEL")
    cols_vides_pm <- intersect(cols_vides_pm, names(pm_stock))

    if (length(cols_vides_pm) > 0) {
      pm_stock_f <- pm_stock %>%
        filter(if_any(all_of(cols_vides_pm), ~ is.na(.) | trimws(.) == ""))
    } else {
      pm_stock_f <- pm_stock[0, ]
    }


  

   if (length(cols_vides_pm) > 0) {
     pm_flux_f <- pm_flux %>%
       filter(if_any(all_of(cols_vides_pm), ~ is.na(.) | trimws(.) == ""))
   } else {
     pm_flux_f <- pm_flux[0, ]
   }


     pm_stock   <- dplyr::mutate(pm_stock,   dplyr::across(dplyr::everything(), clean_spaces))
     pm_stock_f <- dplyr::mutate(pm_stock_f, dplyr::across(dplyr::everything(), clean_spaces))
     pm_flux    <- dplyr::mutate(pm_flux,    dplyr::across(dplyr::everything(), clean_spaces))
     pm_flux_f  <- dplyr::mutate(pm_flux_f,  dplyr::across(dplyr::everything(), clean_spaces))




     readr::write_delim(pm_stock,   paste0(chemin,fil,"//data//pm_",fil,"_STOCK.csv"),
                        delim = ";", quote = "all", escape = "double", na = "")
     readr::write_delim(pm_stock_f, paste0(chemin,fil,"//data//pm_",fil,"_STOCK_F.csv"),
                        delim = ";", quote = "all", escape = "double", na = "")
     readr::write_delim(pm_flux,    paste0(chemin,fil,"//data//pm_",fil,".csv"),
                        delim = ";", quote = "all", escape = "double", na = "")
     readr::write_delim(pm_flux_f,  paste0(chemin,fil,"//data//pm_",fil,"_F.csv"),
                        delim = ";", quote = "all", escape = "double", na = "")

    taux_incomp_pm=paste(round(nrow(pm_stock_f)/nrow(pm_stock)*100,1),"%",sep="")

      taux_incomp_pm_f=paste(round(nrow(pm_flux_f)/nrow(pm_flux)*100,1),"%",sep="")

# ============================
# IMPORT PP_STOCK
# ============================
chemin_pp <- file.path(chemin, fil, "data", "pp_stock.csv")

# Vérification que le fichier existe avant lecture
if (!file.exists(chemin_pp)) {
  stop("Fichier introuvable : ", chemin_pp)
}

# Import du CSV
# read_csv2_big : lecture rapide fread (multi-thread) + cache .rds,
# gère encodage, espaces parasites, guillemets ; fallback read_csv2_auto
pp_stock <- read_csv2_big(chemin_pp, drop = cols_ignorees)
pp_stock=unique(pp_stock)

# Vérification

## Traitement des PP

if (fil=="ML") {

pp_stock_s=pp_stock

pp_stock_s$CLIENT=as.numeric(pp_stock_s$CLIENT)

risque_s=read.csv2(paste(sep="",chemin,"ML//data//risque.csv"),header=F)
risque_s[,1]=as.numeric(risque_s[,1])
risque=data.frame(CLIENT=risque_s[,1], RISQUE=risque_s[,3])

pp_stock_s=left_join(pp_stock_s,risque, by='CLIENT')



pp_stock_s$RISQUE.x = pp_stock_s$RISQUE.y


pp_stock_s=pp_stock_s[,-which(colnames(pp_stock_s)=="RISQUE.y")]

colnames(pp_stock_s)[which(colnames(pp_stock_s)=="RISQUE.x")] = "RISQUE"

pp_stock_s$RESID[pp_stock_s$PAYS_RESID=="MALI"]="O"
pp_stock_s$RESID[pp_stock_s$PAYS_RESID!="MALI"]="N"

pp_stock = pp_stock_s


} else {
   pp_stock=pp_stock
}



nrow(pp_stock)
colnames(pp_stock)

if ("AGENCE_LIB" %in% names(pp_stock)) {
  colnames(pp_stock)[which(names(pp_stock)=="AGENCE_LIB")] <- "AGENCELIB"

} 

if ("LIB_AGENCE" %in% names(pp_stock)) {
   colnames(pp_stock)[which(names(pp_stock)=="LIB_AGENCE")] <- "AGENCELIB"
}

  

    
    pp_stock <- clean_colnames(pp_stock)
    
    pp_stock[is.na(pp_stock)]=""
    
    cli_nul=which(pp_stock$CLIENT=="")
    if (length(cli_nul)!=0) {
        pp_stock=pp_stock[-cli_nul,]
    } else {
        pp_stock=pp_stock
    }
    
     pp_stock=as.data.frame(lapply(pp_stock, remove_spaces))

     if ("AGENCE" %in% names(pp_stock)) {
        pp_stock$AGENCE <- pad_agence5(pp_stock$AGENCE)
     }
    
    pp_stock$DATOUV = dmy(pp_stock$DATOUV)

    diff=setdiff(pp_stock$CLIENT, non_fiab$CLIENT)

    
    diff_sec_pp=intersect(non_fiab$CLIENT, pp_stock$CLIENT)

    pp_stock= pp_stock[pp_stock$CLIENT %in% diff,]
    
    
    clients_devise_pp = which(pp_stock$DEVISE %in% devise_fil)

    pp_stock[clients_devise_pp,"DEVISE"]= ""


    pp_flux <- pp_stock %>%
    filter(
        DATOUV >= as.Date(premier_jour),
        DATOUV < as.Date(jour_recent)
    )


 cols_vides_pp <- c("CODAPE","PAYNAIS","DATNAIS","PROFESSION","ADRESSE","PAYS_RESID",
                    "NUMID","SALAIRE","ORIGINE_REVENU","DATVALID","TEL")
 cols_vides_pp <- intersect(cols_vides_pp, names(pp_stock))

 if (length(cols_vides_pp) > 0) {
   pp_stock_f <- pp_stock %>%
     filter(if_any(all_of(cols_vides_pp), ~ is.na(.) | trimws(.) == ""))
 } else {
   pp_stock_f <- pp_stock[0, ]
 }


  

   if (length(cols_vides_pp) > 0) {
     pp_flux_f <- pp_flux %>%
       filter(if_any(all_of(cols_vides_pp), ~ is.na(.) | trimws(.) == ""))
   } else {
     pp_flux_f <- pp_flux[0, ]
   }


     pp_stock   <- dplyr::mutate(pp_stock,   dplyr::across(dplyr::everything(), clean_spaces))
     pp_flux    <- dplyr::mutate(pp_flux,    dplyr::across(dplyr::everything(), clean_spaces))
     pp_stock_f <- dplyr::mutate(pp_stock_f, dplyr::across(dplyr::everything(), clean_spaces))
     pp_flux_f  <- dplyr::mutate(pp_flux_f,  dplyr::across(dplyr::everything(), clean_spaces))



     readr::write_delim(pp_stock, paste0(chemin,fil,"//data//pp_",fil,"_STOCK.csv"),
                        delim = ";", quote = "all", escape = "double", na = "")
     readr::write_delim(pp_flux, paste0(chemin,fil,"//data//pp_",fil,".csv"),
                        delim = ";", quote = "all", escape = "double", na = "")

     readr::write_delim(pp_stock_f, paste0(chemin,fil,"//data//pp_",fil,"_STOCK_F.csv"),
                        delim = ";", quote = "all", escape = "double", na = "")
     readr::write_delim(pp_flux_f, paste0(chemin,fil,"//data//pp_",fil,"_F.csv"),
                        delim = ";", quote = "all", escape = "double", na = "")

    taux_incomp_pp=paste(round(nrow(pp_stock_f)/nrow(pp_stock)*100,1),"%",sep="")
    taux_incomp_pp_f=paste(round(nrow(pp_flux_f)/nrow(pp_flux)*100,1),"%",sep="")



  #Archivage
          # Liste des fichiers dans le dossier (sans le chemin complet)

      repertoire=file.path(chemin,fil,"Archives")
      dir.create(repertoire)
      fichiers_dossier <- list.files(paste0(chemin,fil), full.names = FALSE)

      # Spécifiez les mots-clés à chercher
      mots_cles <- c("contrôle","suivi","data","Archives")

      # Filtrer les fichiers à conserver
      fichiers_a_conserver <- fichiers_dossier[
        grepl(paste(mots_cles, collapse = "|"), fichiers_dossier)
      ]

      # Fichiers à archiver (sans chemins absolus)
      fichiers_a_archiver <- setdiff(fichiers_dossier, fichiers_a_conserver)

      if (length(fichiers_a_archiver)!=0) {
          repertoire=paste0(chemin,fil,"//",fichiers_a_archiver)
      zipr(paste0(chemin,fil,"//Archives","//Archives ",fil," _ ",premier_jour,".zip"), repertoire)

          unlink(repertoire, recursive = TRUE)
          }


     repertoire=file.path(chemin,"Images")
     dir.create(repertoire)
     sigle=paste("BOA_",fil, sep="")
  
         ## les fichiers excel
     zone=read.csv2(paste(sep="",chemin,"zone_",fil,".csv"), fileEncoding = "UTF-8-BOM")
     zone <- clean_colnames(zone)

     zone$AGENCE=as.numeric(zone$AGENCE)
     # (l'ancien read_pptx de "Temp rapport suivi.pptx" etait inutilise :
     #  l'objet etait ecrase plus bas. Le rapport DG est genere par code.)
    
   
      
            


#####-----------------------DONNEES PAR AGENCE-------------###########
      cat("######################################################\n")
      cat("######### ETAPE 1: TAUX DE FIABILISATION PAR AGENCE \n")
      cat("#####################################################\n")


taux_function_pp=function(x) {

            if (exploitant=="agent") {
                 expl=pp_ne[pp_ne$EXPL==x,]
            } else {
                expl=pp_ne[pp_ne$AGENCE==x,]
            }

            n_expl_s=expl[expl$PAYNAIS %in% inc  | is.na(expl$PAYNAIS)=="TRUE" | expl$PROFESSION %in% inc  | is.na(expl$PROFESSION)=="TRUE"
                    | expl$SALAIRE %in% inc  | is.na(expl$SALAIRE)=="TRUE" | expl$CODAPE %in% inc  | is.na(expl$CODAPE)=="TRUE" 
                    | expl$TEL %in% inc  | is.na(expl$TEL)=="TRUE" | expl$ADRESSE %in% inc  | is.na(expl$ADRESSE) | expl$NUMID %in% inc  | is.na(expl$NUMID)=="TRUE" | expl$DATNAIS %in% inc  | is.na(expl$DATNAIS)=="TRUE" | expl$DATVALID %in% inc  | is.na(expl$DATVALID)=="TRUE" | expl$ORIGINE_REV %in% inc  | is.na(expl$ORIGINE_REV)=="TRUE" | expl$PAYS_RESID %in% inc  | is.na(expl$PAYS_RESID)=="TRUE"  ,]
            n_cons=length(unique(n_expl_s$CLIENT))  
            
                a=field_completion_rate(expl, "PAYNAIS", inc)
                b=field_completion_rate(expl, "PROFESSION", inc)
                c=field_completion_rate(expl, "SALAIRE", inc)
                d=field_completion_rate(expl, "CODAPE", inc)
                e=field_completion_rate(expl, "TEL", inc)
                f=field_completion_rate(expl, "ADRESSE", inc)
                g=field_completion_rate(expl, "NUMID", inc)
                h=field_completion_rate(expl, "DATNAIS", inc)
                i=field_completion_rate(expl, "DATVALID", inc)
                j=field_completion_rate(expl, "ORIGINE_REVENU", inc)
                k=field_completion_rate(expl, "PAYS_RESID", inc)

                taux=min_rate(a,b,c,d,e,f,g,h,i,j,k)
                appreciation=get_appreciation(taux, Faible, Moyen)

                #if (taux_rat<0) {
                #taux_ratt= paste("En avance de + ",abs(taux_rat),"%", sep="")
                #} else {
                #taux_ratt= paste("En retard de ",taux_rat,"%", sep="")
                #}

                taux=format_percent(taux)
                lieu_naiss=format_percent(a)
                profession=format_percent(b)
                revenu=format_percent(c)
                codape=format_percent(d)
                tel=format_percent(e)
                adresse=format_percent(f)
                nin=format_percent(g)
                datnais=format_percent(h)
                datvalid=format_percent(i)
                origine=format_percent(j)
                pays_resid=format_percent(k)


              if (exploitant=="agent") {
                    etat_expls=data.frame(Agents=x,`Nbre clients concernés`=n_cons,`Lieu de Naissance`=lieu_naiss, Profession =profession, Codape=codape, Revenu = revenu, NIN=nin, TEL=tel,Adresse=adresse,DATNAIS=datnais, DATVALID=datvalid, Origine_Revenu=origine, Pays_Residence=pays_resid, `Taux de fiabilisation`=taux, Appréciation=appreciation) # ##, `Taux à rattrapper`=taux_ratt)
            
              } else {
                    etat_expls=data.frame(Agence=x,`Nbre clients concernés`=n_cons,`Lieu de Naissance`=lieu_naiss, Profession =profession, Codape=codape, Revenu = revenu, NIN=nin, TEL=tel,Adresse=adresse,DATNAIS=datnais, DATVALID=datvalid, Origine_Revenu=origine, Pays_Residence=pays_resid, `Taux de fiabilisation`=taux, Appréciation=appreciation) # ##, `Taux à rattrapper`=taux_ratt)

              }

              colnames(etat_expl)=colnames(etat_expls)
              etat_filiale = rbind(etat_expl,etat_expls)
            }


        

        taux_function_pm=function(x) {
            if (exploitant=="agent") {
                 expl=pm_ne[pm_ne$EXPL==x,]
            } else {
                expl=pm_ne[pm_ne$AGENCE==x,]
            }
            
                n_expl=expl[expl$AGEC %in% inc  | is.na(expl$AGEC)=="TRUE" | expl$CAPITAL %in% inc  | is.na(expl$CAPITAL)=="TRUE" | expl$CA %in% inc  | is.na(expl$CA)=="TRUE" | 
                expl$RESULTAT %in% inc  | is.na(expl$RESULTAT)=="TRUE" | expl$RCSNO %in% inc  | is.na(expl$RCSNO)=="TRUE" |
                expl$CODAPE %in% inc  | is.na(expl$CODAPE)=="TRUE" |
                expl$ORIGINE_REV %in% inc  | is.na(expl$ORIGINE_REV)=="TRUE" |
                expl$TEL %in% inc  | is.na(expl$TEL)=="TRUE" ,] 


             
            
            n_cons=length(unique(n_expl$CLIENT))

         
       
            
                    a=field_completion_rate(expl, "CAPITAL", inc)
                    b=field_completion_rate(expl, "CA", inc)
                    c=field_completion_rate(expl, "RESULTAT", inc)
                    d=field_completion_rate(expl, "RCSNO", inc)
                    e=field_completion_rate(expl, "CODAPE", inc)
                    f=field_completion_rate(expl, "AGEC", inc)
                    g=field_completion_rate(expl, "ORIGINE_REVENU", inc)
                    h=field_completion_rate(expl, "TEL", inc)

                    taux=min_rate(a,b,c,d,e,f,g,h)
                    appreciation=get_appreciation(taux, Faible, Moyen)
                    taux_ratt=get_taux_ratt(taux, trimestre)

                    taux=format_percent(taux)
                    capital=format_percent(a)
                    ca=format_percent(b)
                    resultat=format_percent(c)
                    rcsno=format_percent(d)
                    codape=format_percent(e)

                    agec=format_percent(f)
                    origine=format_percent(g)
                    tel=format_percent(h)

                    if (exploitant=="agent") {
                        etat_expls=data.frame(Agents=x,`Nbre clients concernés`=n_cons,AGEC=agec,`Capital`=capital, CA = ca, Resultat = resultat, RCSNO = rcsno,CODAPE= codape, Origine_Revenu=origine, TEL=tel,`Taux de fiabilisation`=taux, Appréciation=appreciation) # ##, `Taux à rattrapper`=taux_ratt)
                    } else {
                        etat_expls=data.frame(Agence=x,`Nbre clients concernés`=n_cons,AGEC=agec, `Capital`=capital, CA = ca, Resultat = resultat, RCSNO = rcsno,CODAPE= codape, Origine_Revenu=origine, TEL=tel, `Taux de fiabilisation`=taux, Appréciation=appreciation) # ##, `Taux à rattrapper`=taux_ratt)
                    }

                    colnames(etat_expls)=colnames(etat_expl)

                    etat_filiale = rbind(etat_expl,etat_expls)
            }
        #function_kyc= function(fil){
        

          
          
          tableau_suivi=data.frame(Flux=c("","Taux de fiabilisation","Appréciation"))
          tableau_suivi_stock=data.frame(Stock=c("","Taux de fiabilisation","Appréciation"))
          
          
          sigle=paste("BOA_",fil, sep="")
          
       
            
            cat("######################################################\n")
            cat("######### Chargement des données de\n",sigle, "#######\n")
            cat("#####################################################\n")
            # Importatations des donnees
          
            # pp_flux/pm_flux sont déjà en mémoire (dérivés de pp_stock/pm_stock),
            # inutile de relire les CSV qui viennent d'être écrits
            pp_ne <- pp_flux
            pp_ne$AGENCE=as.numeric(pp_ne$AGENCE)


            pp_ne=select_pp_fields(pp_ne, pp_fields)

            colnames(pp_ne)=toupper(colnames(pp_ne))
            pm_ne <- pm_flux
            pm_ne$AGENCE=as.numeric(pm_ne$AGENCE)



            pm_ne=select_pm_fields(pm_ne, pm_fields)

            colnames(pm_ne)=toupper(colnames(pm_ne))
                
             pp_ne$EXPL[is_alphanumeric(pp_ne$EXPL)=="FALSE"]=paste("DIR_AGENCE_",pp_ne$AGENCE[is_alphanumeric(pp_ne$EXPL)=="FALSE"],sep="")
             pm_ne$EXPL[is_alphanumeric(pm_ne$EXPL)=="FALSE"]=paste("DIR_AGENCE_",pm_ne$AGENCE[is_alphanumeric(pm_ne$EXPL)=="FALSE"],sep="")





            if (exists("pp_ne")=="TRUE" & exists("pm_ne")=="TRUE") {
              
              cat("########################################################################\n")
              cat("######### Les données de\n",sigle, "ont été chargées avec succès #######\n")
              cat("########################################################################\n")
              
            } else {
              
              cat("################################################################################################\n")
              cat("######### Erreur: les données de\n",sigle, "n'ont pas été bien chargées (Voir les formats) #####\n")
              cat("###############################################################################################\n")
              
            }
            
            
            
            
            ## Appreciation flux
            
            trimestre=100
            Faible=80
            Moyen=100
            Bon=100
            
            ## Statistique a l'échelle filiale (PP)
       
        a=field_completion_rate(pp_ne, "PAYNAIS", inc)
        b=field_completion_rate(pp_ne, "PROFESSION", inc)
        c=field_completion_rate(pp_ne, "SALAIRE", inc)
        d=field_completion_rate(pp_ne, "CODAPE", inc)
        e=field_completion_rate(pp_ne, "TEL", inc)
        f=field_completion_rate(pp_ne, "ADRESSE", inc)
        g=field_completion_rate(pp_ne, "NUMID", inc)
        h=field_completion_rate(pp_ne, "DATNAIS", inc)
        i=field_completion_rate(pp_ne, "DATVALID", inc)
        j=field_completion_rate(pp_ne, "ORIGINE_REVENU", inc)
        k=field_completion_rate(pp_ne, "PAYS_RESID", inc)
        


        T=length(c(a,b,c,d,e,f,g,h,i,j))
        
       
                    
      #  taux=floor(100*(1-sum(nrow(pp_ne[pp_ne$PAYNAIS=="" | is.na(pp_ne$PAYNAIS)=="TRUE",]),nrow(pp_ne[pp_ne$PROFESSION=="" | is.na(pp_ne$PROFESSION)=="TRUE",]),
      #           nrow(pp_ne[pp_ne$SALAIRE=="" | is.na(pp_ne$SALAIRE)=="TRUE",]), nrow(pp_ne[pp_ne$NUMID=="" | is.na(pp_ne$NUMID)=="TRUE",]),
      #           nrow(pp_ne[pp_ne$CODAPE=="" | is.na(pp_ne$CODAPE)=="TRUE",]),nrow(pp_ne[pp_ne$TEL=="" | is.na(pp_ne$TEL)=="TRUE",]),          
      #           nrow(pp_ne[pp_ne$DATNAIS=="" | is.na(pp_ne$DATNAIS)=="TRUE",]),
      #           nrow(pp_ne[pp_ne$ADRESSE=="" |is.na(pp_ne$ADRESSE)=="TRUE",]),   nrow(pp_ne[pp_ne$NUMID=="" |is.na(pp_ne$NUMID)=="TRUE",]), nrow(pp_ne[pp_ne$DATVALID=="" |is.na(pp_ne$DATVALID)=="TRUE" ,]), 
      #                     nrow(pp_ne[pp_ne$ORIGINE_REV=="" |is.na(pp_ne$ORIGINE_REV)=="TRUE" ,])
      #           
      #           )/(T*nrow(pp_ne))))
      # 

       
       
        champs=c("PAYNAIS","PROFESSION","SALAIRE","NUMID","CODAPE","TEL","DATNAIS","ADRESSE","DATVALID","ORIGINE_REVENU","PAYS_RESID")

        taux=compute_taux(pp_ne, champs)
        appreciation=get_appreciation(taux, Faible, Moyen)
        taux_ratt=get_taux_ratt(taux, trimestre)
        taux=format_percent(taux)

        lieu_naiss=format_percent(a)
        profession=format_percent(b)
        revenu=format_percent(c)
        codape=format_percent(d)
        tel=format_percent(e)
        adresse=format_percent(f)
        nin=format_percent(g)
        datnais=format_percent(h)
        datvalid=format_percent(i)
        origine=format_percent(j)
        pays_resid=format_percent(k)



       
        etat_pp_nes=data.frame(`Lieu de Naissance` = lieu_naiss, Profession = profession,Codape=codape, Revenu = revenu, NIN = nin, TEL=tel, Adresse=adresse,DATNAIS=datnais,DATVALID=datvalid, ORIGINE_REV=origine, PAYS_RESID=pays_resid, `Taux de fiabilisation` =taux, Appréciation=appreciation, `Taux à rattrapper` = taux_ratt)
            pp=etat_pp_nes[,c(ncol(etat_pp_nes),ncol(etat_pp_nes)-1,ncol(etat_pp_nes)-2)]
            pp_t=data.frame(A="",V="")
            
            colnames(pp_t)=c(paste(fil),"")
            pp_t[1,]=c("PM","PP")
            pp_t[2,2]=etat_pp_nes$Taux.de.fiabilisation[1]
            pp_t[3,2]=etat_pp_nes$Appréciation[1]
            
            rownames(pp_t)=c("","Taux de fiabilisation","Appréciation")
            ## Statistique a l'échelle filiale (PM)
            

            
            a=field_completion_rate(pm_ne, "CAPITAL", inc)
            b=field_completion_rate(pm_ne, "CA", inc)
            c=field_completion_rate(pm_ne, "RESULTAT", inc)
            d=field_completion_rate(pm_ne, "RCSNO", inc)
            e=field_completion_rate(pm_ne, "CODAPE", inc)
            f=field_completion_rate(pm_ne, "AGEC", inc)
            g=field_completion_rate(pm_ne, "ORIGINE_REVENU", inc)
            h=field_completion_rate(pm_ne, "TEL", inc)

           
            K=length(c(a,b,c,d,e,f,g,h))

         #   taux=floor(100*(1-sum(nrow(pm_ne[pm_ne$AGEC %in% inc | is.na(pm_ne$AGEC)=="TRUE",]), nrow(pm_ne[pm_ne$CAPITAL %in% inc | is.na(pm_ne$CAPITAL)=="TRUE",]), nrow(pm_ne[pm_ne$CA %in% inc | is.na(pm_ne$CA)=="TRUE",]),
         #
              #nrow(pm_ne[pm_ne$RESULTAT %in% inc | is.na(pm_ne$RESULTAT)=="TRUE",]), nrow(pm_ne[pm_ne$RCSNO %in% inc | is.na(pm_ne$RCSNO=="TRUE"),]),
         #
             # nrow(pm_ne[pm_ne$CODAPE %in% inc | is.na(pm_ne$CODAPE=="TRUE"),]), nrow(pm_ne[pm_ne$ORIGINE_REV %in% inc | is.na(pm_ne$ORIGINE_REV)=="TRUE",]), nrow(pm_ne[pm_ne$TEL %in% inc | is.na(pm_ne$TEL)=="TRUE",]))/(K*nrow(pm_ne))))
#

        champs=c("CAPITAL","CA","RESULTAT","RCSNO","CODAPE","AGEC","ORIGINE_REVENU","TEL")

        taux=compute_taux(pm_ne, champs)
        appreciation=get_appreciation(taux, Faible, Moyen)
        taux_ratt=get_taux_ratt(taux, trimestre, label_retard = FALSE)
        taux=format_percent(taux)
        capital=format_percent(a)
        ca=format_percent(b)
        resultat=format_percent(c)
        rcsno=format_percent(d)
        codape=format_percent(e)
        agec=format_percent(f)
        origine=format_percent(g)
        tel=format_percent(h)


            
            etat_pm_nes=data.frame(AGEC=agec,`Capital`=capital, CA = ca, Resultat = resultat, RCSNO = rcsno,CODAPE= codape,  ORIGINE_REV=origine, TEL=tel, `Taux de fiabilisation`=taux, Appréciation=appreciation) # ##, `Taux à rattrapper`=taux_ratt)
            
            
            pm=etat_pm_nes[,c(ncol(etat_pm_nes),ncol(etat_pm_nes)-1,ncol(etat_pm_nes)-2)]
            
            pp_t[2,1]=c(pm[1,c(2)])
            pp_t[3,1]=c(pm[1,c(1)])

            tableau_suivi = cbind(tableau_suivi,pp_t)
            tableau_suivi[,1]=rownames(tableau_suivi)
            
            ## PP
            
            agents=unique(pp_ne$AGENCE)#("[[:alnum:]]", pp_ne$EXPL)=="TRUE"])
            #agents_null_pp=pp_ne[grepl("[[:alnum:]]", pp_ne$EXPL)=="FALSE",]

            exploitant="agence"
            

            etat_expl=data.frame(Agence="",`Nbre clients concernés`="",`Lieu de Naissance`="", Profession = "",Codape="", Revenu = "", NIN="", Tel="",Adresse="",DATNAIS="", DATVALID="", ORIGINE_REV="", PAYS_RESID="", `Taux de fiabilisation`="", Appréciation="") # ##, `Taux à rattrapper`="")
             taux_filiale = lapply(agents,taux_function_pp)
            
            taux_filiale_t=bind_results(taux_filiale, etat_expl)
            taux_filiale_t = drop_lignes_vides(taux_filiale_t, "NIN")
            
      
     
            ## PM
            
            agents_pm=unique(pm_ne$AGENCE)#("[[:alnum:]]", pm_ne$EXPL)=="TRUE"])
            #agents_null_pm=pm_ne[grepl("[[:alnum:]]", pm_ne$EXPL)=="FALSE",]
            
          
            
            etat_expl=data.frame(Agence="",`Nbre clients concernés`="", AGEC="", Capital="", CA = "", Resultat = "", RCSNO = "", CODAPE= "",  ORIGINE_REV=origine, TEL=tel,  `Taux de fiabilisation`="", Appréciation="")#,  `Taux à rattrapper`="")

            taux_filiale_pm = lapply(agents_pm,taux_function_pm)
            
            taux_filiale_pm=bind_results(taux_filiale_pm, etat_expl)
            taux_filiale_pm = drop_lignes_vides(taux_filiale_pm, "RCSNO")
            
         
               
            
            rm(pp_ne,pm_ne)
            
            
            ### LE stock
            
            cat("######################################################\n")
            cat("######### Chargement des données Stock de\n",sigle, "#######\n")
            cat("#####################################################\n")
            # Importatations des donnees

            # pp_stock/pm_stock sont déjà en mémoire et nettoyés,
            # inutile de relire les CSV _STOCK qui viennent d'être écrits
            pp_ne <- pp_stock
            pp_ne$AGENCE=as.numeric(pp_ne$AGENCE)




            pp_ne=select_pp_fields(pp_ne, pp_fields)

            colnames(pp_ne)=toupper(colnames(pp_ne))
            pm_ne <- pm_stock
            pm_ne$AGENCE=as.numeric(pm_ne$AGENCE)


            pm_ne=select_pm_fields(pm_ne, pm_fields)
            colnames(pm_ne)=toupper(colnames(pm_ne))


                
            pp_ne$EXPL[is_alphanumeric(pp_ne$EXPL)=="FALSE"]=paste("DIR_AGENCE_",pp_ne$AGENCE[is_alphanumeric(pp_ne$EXPL)=="FALSE"],sep="")
            pm_ne$EXPL[is_alphanumeric(pm_ne$EXPL)=="FALSE"]=paste("DIR_AGENCE_",pm_ne$AGENCE[is_alphanumeric(pm_ne$EXPL)=="FALSE"],sep="")




            if (exists("pp_ne")=="TRUE" & exists("pm_ne")=="TRUE") {
              
              cat("########################################################################\n")
              cat("######### Les données Stock de\n",sigle, "ont été chargées avec succès #######\n")
              cat("########################################################################\n")
              
            } else {
              
              cat("################################################################################################\n")
              cat("######### Erreur: les données Stock de\n",sigle, "n'ont pas été bien chargées (Voir les formats) #####\n")
              cat("###############################################################################################\n")
              
            }


            
            
            ##-----Appréciation du Stock-####
            
            if (trimestre_actuel=="1"){
              trimestre=30
              Faible=5
              Moyen=30
              
            } else if (trimestre_actuel=="2") {
              trimestre=60
              Faible=30
              Moyen=60
            }  else if (trimestre_actuel=="3") {
              trimestre=90
              Faible=60
              Moyen=90
            }  else {
              trimestre=95
              Faible=65
              Moyen=95
            }
            
            
            
            ## Statistique a l'échelle filiale (PP)
            
        a=field_completion_rate(pp_ne, "PAYNAIS", inc)
        b=field_completion_rate(pp_ne, "PROFESSION", inc)
        c=field_completion_rate(pp_ne, "SALAIRE", inc)
        d=field_completion_rate(pp_ne, "CODAPE", inc)
        e=field_completion_rate(pp_ne, "TEL", inc)
        f=field_completion_rate(pp_ne, "ADRESSE", inc)
        g=field_completion_rate(pp_ne, "NUMID", inc)
        h=field_completion_rate(pp_ne, "DATNAIS", inc)
        i=field_completion_rate(pp_ne, "DATVALID", inc)
        j=field_completion_rate(pp_ne, "ORIGINE_REVENU", inc)
        k=field_completion_rate(pp_ne, "PAYS_RESID", inc)

        T=length(c(a,b,c,d,e,f,g,h,i,j))
        
       
               
       #taux=floor(100*(1-sum(nrow(pp_ne[pp_ne$PAYNAIS=="" | is.na(pp_ne$PAYNAIS)=="TRUE",]),nrow(pp_ne[pp_ne$PROFESSION=="" | is.na(pp_ne$PROFESSION)=="TRUE",]),
       #         nrow(pp_ne[pp_ne$SALAIRE=="" | is.na(pp_ne$SALAIRE)=="TRUE",]), nrow(pp_ne[pp_ne$NUMID=="" | is.na(pp_ne$NUMID)=="TRUE",]),
       #         nrow(pp_ne[pp_ne$CODAPE=="" | is.na(pp_ne$CODAPE)=="TRUE",]),nrow(pp_ne[pp_ne$TEL=="" | is.na(pp_ne$TEL)=="TRUE",]),          
       #         nrow(pp_ne[pp_ne$DATNAIS=="" | is.na(pp_ne$DATNAIS)=="TRUE",]),
       #         nrow(pp_ne[pp_ne$ADRESSE=="" |is.na(pp_ne$ADRESSE)=="TRUE",]),   nrow(pp_ne[pp_ne$NUMID=="" |is.na(pp_ne$NUMID)=="TRUE",]), nrow(pp_ne[pp_ne$DATVALID=="" |is.na(pp_ne$DATVALID)=="TRUE",]), 
       #                   nrow(pp_ne[pp_ne$ORIGINE_REV=="" |is.na(pp_ne$ORIGINE_REV)=="TRUE",])
       #         
       #         )/(T*nrow(pp_ne))))
       #

       
       
        champs=c("PAYNAIS","PROFESSION","SALAIRE","NUMID","CODAPE","TEL","DATNAIS","ADRESSE","DATVALID","ORIGINE_REVENU","PAYS_RESID")

        taux=compute_taux(pp_ne, champs)
        appreciation=get_appreciation(taux, Faible, Moyen)
        taux_ratt=get_taux_ratt(taux, trimestre)
        taux=format_percent(taux)

        lieu_naiss=format_percent(a)
        profession=format_percent(b)
        revenu=format_percent(c)
        codape=format_percent(d)
        tel=format_percent(e)
        adresse=format_percent(f)
        nin=format_percent(g)
        datnais=format_percent(h)
        datvalid=format_percent(i)
        origine=format_percent(j)
        pays_resid=format_percent(k)



        etat_pp_nes=data.frame(`Lieu de Naissance` = lieu_naiss, Profession = profession,Codape=codape, Revenu = revenu, NIN = nin, TEL=tel, Adresse=adresse,DATNAIS=datnais, DATVALID=datvalid,ORIGINE_REV=origine, PAYS_RESID=pays_resid, `Taux de fiabilisation` =taux, Appréciation=appreciation, `Taux à rattrapper` = taux_ratt)
        pp=etat_pp_nes[,c(ncol(etat_pp_nes),ncol(etat_pp_nes)-1,ncol(etat_pp_nes)-2)]
        pp_t=data.frame(A="",V="")
            
            colnames(pp_t)=c(paste(fil),"")
            pp_t[1,]=c("PM","PP")
            pp_t[2,2]=etat_pp_nes$Taux.de.fiabilisation[1]
            pp_t[3,2]=etat_pp_nes$Appréciation[1]
            
            rownames(pp_t)=c("","Taux de fiabilisation","Appréciation")
            
            ## Statistique a l'échelle filiale (PM)
                     
            a=field_completion_rate(pm_ne, "CAPITAL", inc)
            b=field_completion_rate(pm_ne, "CA", inc)
            c=field_completion_rate(pm_ne, "RESULTAT", inc)
            d=field_completion_rate(pm_ne, "RCSNO", inc)
            e=field_completion_rate(pm_ne, "CODAPE", inc)
            f=field_completion_rate(pm_ne, "AGEC", inc)
            g=field_completion_rate(pm_ne, "ORIGINE_REVENU", inc)
            h=field_completion_rate(pm_ne, "TEL", inc)

           
            K=length(c(a,b,c,d,e,f,g,h))

          #  taux=floor(100*(1-sum(nrow(pm_ne[pm_ne$AGEC %in% inc | is.na(pm_ne$AGEC)=="TRUE",]), nrow(pm_ne[pm_ne$CAPITAL %in% inc | is.na(pm_ne$CAPITAL)=="TRUE",]), nrow(pm_ne[pm_ne$CA %in% inc | is.na(pm_ne$CA)=="TRUE",]),
          #    nrow(pm_ne[pm_ne$RESULTAT %in% inc | is.na(pm_ne$RESULTAT)=="TRUE",]), nrow(pm_ne[pm_ne$RCSNO %in% inc | is.na(pm_ne$RCSNO=="TRUE"),]),
          #    nrow(pm_ne[pm_ne$CODAPE %in% inc | is.na(pm_ne$CODAPE=="TRUE"),]), nrow(pm_ne[pm_ne$ORIGINE_REV %in% inc | is.na(pm_ne$ORIGINE_REV)=="TRUE",]), nrow(pm_ne[pm_ne$TEL %in% inc | is.na(pm_ne$TEL)=="TRUE",]))/(K*nrow(pm_ne))))
#
        champs=c("CAPITAL","CA","RESULTAT","RCSNO","CODAPE","AGEC","ORIGINE_REVENU","TEL")

        taux=compute_taux(pm_ne, champs)
        appreciation=get_appreciation(taux, Faible, Moyen)
        taux_ratt=get_taux_ratt(taux, trimestre, label_retard = FALSE)
        taux=format_percent(taux)
        capital=format_percent(a)
        ca=format_percent(b)
        resultat=format_percent(c)
        rcsno=format_percent(d)
        codape=format_percent(e)
        agec=format_percent(f)
        origine=format_percent(g)
        tel=format_percent(h)


            
            etat_pm_nes=data.frame(AGEC=agec,`Capital`=capital, CA = ca, Resultat = resultat, RCSNO = rcsno,CODAPE= codape,  ORIGINE_REV=origine, TEL=tel, `Taux de fiabilisation`=taux, Appréciation=appreciation) # ##, `Taux à rattrapper`=taux_ratt)
            
            
            pm=etat_pm_nes[,c(ncol(etat_pm_nes),ncol(etat_pm_nes)-1,ncol(etat_pm_nes)-2)]
            
           pp_t[2,1]=c(pm[1,c(2)])
            pp_t[3,1]=c(pm[1,c(1)])
            
            tableau_suivi_stock = cbind(tableau_suivi_stock,pp_t)
            tableau_suivi_stock[,1]=rownames(tableau_suivi_stock)
            
            tableau_fiabilisation_stock=cbind(tableau_fiabilisation_stock, tableau_suivi_stock)
            
            if (exists("tableau_fiabilisation_stock")=="TRUE") {
              cat("######################################################################################\n")
              cat("######### Les fichiers Stock des suivi de ",sigle," par agence sont générés avec sucès ######\n")
              cat("######################################################################################\n")
            } else {
              cat("######################################################################################\n")
              cat("######### ERREUR: Les fichiers Stock de suivi de ",sigle," par agence ne sont pas générés ####\n")
              cat("######################################################################################\n")
              
            }
            
            exploitant="agence"
            
            ## PP
            
            agents=unique(pp_ne$AGENCE)            
            
            etat_expl=data.frame(Agence="",`Nbre clients concernés`="", `Lieu de Naissance`= "", Profession = "",Codape="", Revenu = "", NIN= "",  Tel="" ,Adresse="",DATNAIS="", DATVALID="",  ORIGINE_REV="", PAYS_RESID="",`Taux de fiabilisation`="", Appréciation="") # ##, `Taux à rattrapper`="")
             taux_filiale = lapply(agents,taux_function_pp)

            taux_filiale_t_stock=bind_results(taux_filiale, etat_expl)
            taux_filiale_t_stock = drop_lignes_vides(taux_filiale_t_stock, "NIN")
  
            
            
   
            
            ## PM
            
            agents_pm=unique(pm_ne$AGENCE)#("[[:alnum:]]", pm_ne$EXPL)=="TRUE"])
            #agents_null_pm=pm_ne[grepl("[[:alnum:]]", pm_ne$EXPL)=="FALSE",]
            

            etat_expl=data.frame(Agence="", `Nbre clients concernés`="",AGEC="", Capital="", CA = "", Resultat = "", RCSNO = "",CODAPE= "", ORIGINE_REV="", TEL=tel,  `Taux de fiabilisation`="", Appréciation="")#,  `Taux à rattrapper`="")
            taux_filiale_pm_stock = lapply(agents_pm,taux_function_pm)
            
            taux_filiale_pm_stock=bind_results(taux_filiale_pm_stock, etat_expl)
            taux_filiale_pm_stock = drop_lignes_vides(taux_filiale_pm_stock, "RCSNO")
            
      
            
            
            ### Creation du fichier Excel de suivi
            wb=createWorkbook()
            
    
            
            ### Resume filiale
            addWorksheet(wb,"Récapitulatif")
            
            
            k=ncol(tableau_suivi)
    
            ### recap stock
            colnames(tableau_suivi_stock)=c("Stock",fil,"")
            tableau_suivi_stock_t=tableau_suivi_stock[-3,]
            writeData(wb,"Récapitulatif", x=tableau_suivi_stock_t, startRow=5, startCol=3)
            addStyle(wb,"Récapitulatif",headerStyle,cols=3:5, rows=5)
            
            ### recap flux
            colnames(tableau_suivi)=c("Flux",fil,"")
            tableau_suivi_t=tableau_suivi[-3,]
            writeData(wb,"Récapitulatif", x=tableau_suivi_t, startRow=5, startCol=7)
            addStyle(wb,"Récapitulatif",headerStyle,cols=7:9, rows=5)
  
            
            
            
            ### Sheet Flux PP
            addWorksheet(wb,"Flux PP")
            
            
            writeData(wb,"Flux PP", x=taux_filiale_t, startRow=1, startCol=1)
            
            k=ncol(taux_filiale_t)
            
           
            addStyle(wb,"Flux PP",headerStyle,cols=1:k, rows=1)
            

                             
            apply_status_style(wb, "Flux PP", taux_filiale_t, value_col = 15)
            
            ### Sheet Flux PM
            addWorksheet(wb,"Flux PM")
            
            writeData(wb,"Flux PM", x=taux_filiale_pm, startRow=1, startCol=1)
            
            k=ncol(taux_filiale_pm)
            
           
            
            addStyle(wb,"Flux PM",headerStyle,cols=1:k, rows=1)
            
            
            apply_status_style(wb, "Flux PM", taux_filiale_pm, value_col = 12)
            
            
            
            ### Sheet Stock PP
            addWorksheet(wb,"Stock PP")
            
            
            writeData(wb,"Stock PP", x=taux_filiale_t_stock, startRow=1, startCol=1)
            
            k=ncol(taux_filiale_t_stock)
            
         
            
            addStyle(wb,"Stock PP",headerStyle,cols=1:k, rows=1)
            

            apply_status_style(wb, "Stock PP", taux_filiale_t_stock, value_col = 15)
            
            ### Sheet Stock PM
            addWorksheet(wb,"Stock PM")
            
            writeData(wb,"Stock PM", x=taux_filiale_pm_stock, startRow=1, startCol=1)
            
            k=ncol(taux_filiale_pm_stock)
            
       
            
            addStyle(wb,"Stock PM",headerStyle,cols=1:k, rows=1)
            
            
            apply_status_style(wb, "Stock PM", taux_filiale_pm_stock, value_col = 12)
            
            saveWorkbook(wb, paste(sep="",paste(chemin,fil,"//",sep=""),"Rapport du taux de complétude par agence ",sigle,".xlsx"), overwrite=T)
    
            wb_recap=wb
            
            
            tableau_fiabilisation=cbind(tableau_fiabilisation,tableau_suivi)

            #### Les tabmeaux de fiabilisation en vertical

            fiabilisation=cbind(tableau_suivi[2,],tableau_suivi_stock[2,-1])
            fiabilisation[1,1]=fil

            colnames(fiabilisation)[c(1,2:5)]=c("",seq(1,4))
            colnames(fiabilisation)[c(2,4)]=c("Flux","Stock")

            fiabilisation_fil=rbind(fiabilisation_fil,fiabilisation)
            colnames(fiabilisation_fil)[c(2,4)]=c("Flux","Stock")
            
            if (exists("tableau_fiabilisation")=="TRUE") {
              cat("######################################################################################\n")
              cat("######### Les fichiers des suivi de ",sigle," par agence sont générés avec sucès ######\n")
              cat("######################################################################################\n")
            } else {
              cat("######################################################################################\n")
              cat("######### ERREUR: Les fichiers des suivi de ",sigle," par agence ne sont pas générés ####\n")
              cat("######################################################################################\n")
              
            }
 
       

          #####-----------------------IMPORTATION DES LIBRAIRIES-------------###########
      
      cat("###############################################\n")
      cat("######### FIABILISATION DES DONNEES KYC \n")
      cat("######### ETAPE 2: FIABILISATION PAR AGENT\n")
      cat("##############################################\n")




if (file.exists(paste0(chemin,fil,"//suivi_fiabilisation.txt"))) {
    suivi_fiabilisation=read.csv2(paste0(chemin,fil,"//suivi_fiabilisation.txt"))
      if (colnames(suivi_fiabilisation)[1]=="X"){
      suivi_fiabilisation=suivi_fiabilisation[,-1]
  }
} else {
   suivi_fiabilisation=c()
}


    suivi_fiabilisation=as.data.frame(lapply(suivi_fiabilisation, remove_spaces))










    
   repertoire=file.path(chemin,fil,"Contrôle de qualité")
   dir.create(repertoire)

   repertoire=file.path(chemin,fil,"Contrôle de qualité","Contrôle qualité Flux")
   dir.create(repertoire)

   repertoire=file.path(chemin,fil,"Contrôle de qualité","Contrôle qualité Stock")
   dir.create(repertoire)



   tableau_suivi=data.frame(Flux=c("","Taux de fiabilisation","Appréciation"))
   tableau_suivi_stock=data.frame(Stock=c("","Taux de fiabilisation","Appréciation"))
  

    sigle=paste("BOA_",fil, sep="")
 

            cat("######### Chargement des données de\n",sigle, "\n")

        # pp_flux/pm_flux déjà en mémoire, pas de relecture des CSV
        pp_ne <- pp_flux
        pp_ne$AGENCE=as.numeric(pp_ne$AGENCE)

        pp_ne=pp_ne[is.na(pp_ne$AGENCE)=="FALSE",]

        pp_ne=select_pp_fields(pp_ne, pp_fields)
        colnames(pp_ne)=toupper(colnames(pp_ne))

        pm_ne <- pm_flux
          pm_ne$AGENCE=as.numeric(pm_ne$AGENCE)

                pm_ne=pm_ne[is.na(pm_ne$AGENCE)=="FALSE",]


        pm_ne=select_pm_fields(pm_ne, pm_fields)

        colnames(pm_ne)=toupper(colnames(pm_ne))
        
        pp_ne$EXPL[is_alphanumeric(pp_ne$EXPL)=="FALSE"]=paste("DIR_AGENCE_",pp_ne$AGENCE[is_alphanumeric(pp_ne$EXPL)=="FALSE"],sep="")
        pm_ne$EXPL[is_alphanumeric(pm_ne$EXPL)=="FALSE"]=paste("DIR_AGENCE_",pm_ne$AGENCE[is_alphanumeric(pm_ne$EXPL)=="FALSE"],sep="")





        if (exists("pp_ne")=="TRUE" & exists("pm_ne")=="TRUE") {

            cat("######### Les données de\n",sigle, "ont été chargées avec succès \n")

        } else {

            cat("######### Erreur: les données de\n",sigle, "n'ont pas été bien chargées (Voir les formats) #####\n")

        }

pp_ne_f=pp_ne
pm_ne_f=pm_ne

        
        trimestre=100
        Faible=80
        Moyen=100
        Bon=100

        ## Statistique a l'échelle filiale (PP)

     
        a=field_completion_rate(pp_ne, "PAYNAIS", inc)
        b=field_completion_rate(pp_ne, "PROFESSION", inc)
        c=field_completion_rate(pp_ne, "SALAIRE", inc)
        d=field_completion_rate(pp_ne, "CODAPE", inc)
        e=field_completion_rate(pp_ne, "TEL", inc)
        f=field_completion_rate(pp_ne, "ADRESSE", inc)
        g=field_completion_rate(pp_ne, "NUMID", inc)
        h=field_completion_rate(pp_ne, "DATNAIS", inc)
        i=field_completion_rate(pp_ne, "DATVALID", inc)
        j=field_completion_rate(pp_ne, "ORIGINE_REVENU", inc)
        k=field_completion_rate(pp_ne, "PAYS_RESID", inc)

        T=length(c(a,b,c,d,e,f,g,h,i,j))
        
       
             
      #           
      #   taux=floor(100*(1-sum(nrow(pp_ne[pp_ne$PAYNAIS=="" | is.na(pp_ne$PAYNAIS)=="TRUE",]),nrow(pp_ne[pp_ne$PROFESSION=="" | is.na(pp_ne$PROFESSION)=="TRUE",]),
      #            nrow(pp_ne[pp_ne$SALAIRE=="" | is.na(pp_ne$SALAIRE)=="TRUE",]), nrow(pp_ne[pp_ne$NUMID=="" | is.na(pp_ne$NUMID)=="TRUE",]),
      #            nrow(pp_ne[pp_ne$CODAPE=="" | is.na(pp_ne$CODAPE)=="TRUE",]),nrow(pp_ne[pp_ne$TEL=="" | is.na(pp_ne$TEL)=="TRUE",]),          
      #            nrow(pp_ne[pp_ne$DATNAIS=="" | is.na(pp_ne$DATNAIS)=="TRUE",]),
      #            nrow(pp_ne[pp_ne$ADRESSE=="" |is.na(pp_ne$ADRESSE)=="TRUE",]),   nrow(pp_ne[pp_ne$NUMID=="" |is.na(pp_ne$NUMID)=="TRUE",]), nrow(pp_ne[pp_ne$DATVALID=="" |is.na(pp_ne$DATVALID)=="TRUE" ,]), 
      #                      nrow(pp_ne[pp_ne$ORIGINE_REV=="" |is.na(pp_ne$ORIGINE_REV)=="TRUE" ,] )
      #            
      #            )/(T*nrow(pp_ne))))
      #  
      #  

      
       
        champs=c("PAYNAIS","PROFESSION","SALAIRE","NUMID","CODAPE","TEL","DATNAIS","ADRESSE","DATVALID","ORIGINE_REVENU","PAYS_RESID")

        taux=compute_taux(pp_ne, champs)
        appreciation=get_appreciation(taux, Faible, Moyen)
        taux=format_percent(taux)

        lieu_naiss=format_percent(a)
        profession=format_percent(b)
        revenu=format_percent(c)
        codape=format_percent(d)
        tel=format_percent(e)
        adresse=format_percent(f)
        nin=format_percent(g)
        datnais=format_percent(h)
        datvalid=format_percent(i)
        origine=format_percent(j)
        pays_resid=format_percent(k)

       
        etat_pp_nes=data.frame(`Lieu de Naissance` = lieu_naiss, Profession = profession,Codape=codape, Revenu = revenu, NIN = nin, TEL=tel, Adresse=adresse,DATNAIS=datnais,DATVALID=datvalid, ORIGINE_REV=origine, PAYS_RESID=pays_resid, `Taux de fiabilisation` =taux, Appréciation=appreciation)##`Taux à rattrapper` = taux_ratt)
        pp=etat_pp_nes[,c(ncol(etat_pp_nes),ncol(etat_pp_nes)-1,ncol(etat_pp_nes)-2)]
        pp_t=data.frame(A="",V="")

        colnames(pp_t)=c(paste(fil),"")
        pp_t[1,]=c("PM","PP")
        pp_t[2,2]=etat_pp_nes$Taux.de.fiabilisation[1]
        pp_t[3,2]=etat_pp_nes$Appréciation[1]

        rownames(pp_t)=c("","Taux de fiabilisation","Appréciation")
        ## Statistique a l'échelle filiale (PM)

if (nrow(pm_ne)!=0) {


            a=field_completion_rate(pm_ne, "CAPITAL", inc)
            b=field_completion_rate(pm_ne, "CA", inc)
            c=field_completion_rate(pm_ne, "RESULTAT", inc)
            d=field_completion_rate(pm_ne, "RCSNO", inc)
            e=field_completion_rate(pm_ne, "CODAPE", inc)
            f=field_completion_rate(pm_ne, "AGEC", inc)
            g=field_completion_rate(pm_ne, "ORIGINE_REVENU", inc)
            h=field_completion_rate(pm_ne, "TEL", inc)

           
            K=length(c(a,b,c,d,e))

           # taux=floor(100*(1-sum(nrow(pm_ne[pm_ne$AGEC %in% inc | is.na(pm_ne$AGEC)=="TRUE",]), nrow(pm_ne[pm_ne$CAPITAL %in% inc | is.na(pm_ne$CAPITAL)=="TRUE",]), nrow(pm_ne[pm_ne$CA %in% inc | is.na(pm_ne$CA)=="TRUE",]),
           #   nrow(pm_ne[pm_ne$RESULTAT %in% inc | is.na(pm_ne$RESULTAT)=="TRUE",]), nrow(pm_ne[pm_ne$RCSNO %in% inc | is.na(pm_ne$RCSNO=="TRUE"),]),
           #   nrow(pm_ne[pm_ne$CODAPE %in% inc | is.na(pm_ne$CODAPE=="TRUE"),]), nrow(pm_ne[pm_ne$ORIGINE_REV %in% inc | is.na(pm_ne$ORIGINE_REV)=="TRUE",]), nrow(pm_ne[pm_ne$TEL %in% inc | is.na(pm_ne$TEL)=="TRUE",]))/(K*nrow(pm_ne))))
#
       champs=c("CAPITAL","CA","RESULTAT","RCSNO","CODAPE","AGEC","ORIGINE_REVENU","TEL")

        taux=compute_taux(pm_ne, champs)
        appreciation=get_appreciation(taux, Faible, Moyen)
        taux=format_percent(taux)
        capital=format_percent(a)
        ca=format_percent(b)
        resultat=format_percent(c)
        rcsno=format_percent(d)
        codape=format_percent(e)
        agec=format_percent(f)
        origine=format_percent(g)
        tel=format_percent(h)


            
            etat_pm_nes=data.frame(AGEC=agec,`Capital`=capital, CA = ca, Resultat = resultat, RCSNO = rcsno,CODAPE= codape,  ORIGINE_REV=origine, TEL=tel, `Taux de fiabilisation`=taux, Appréciation=appreciation)##`Taux à rattrapper`=taux_ratt)
            
        pm=etat_pm_nes[,c(ncol(etat_pm_nes),ncol(etat_pm_nes)-1,ncol(etat_pm_nes)-2)]

       pp_t[2,1]=c(pm[1,c(2)])
            pp_t[3,1]=c(pm[1,c(1)])

          } else {
                    pp_t=data.frame(MR=c("PM","N/A","Aucune donnée"),RM=c("PP","N/A","Aucune donnée"))
                    colnames(pp_t)=c(fil,"")
                    rownames(pp_t)=c("","Taux de fiabilisation","Appréciation")
                  }


        tableau_suivi = cbind(tableau_suivi,pp_t)

        tableau_suivi = tableau_suivi[,-1]

   
        

        ## PP
        
        agents=unique(pp_ne$EXPL[grepl("[[:alnum:]]", pp_ne$EXPL)=="TRUE"])
        agents_null_pp=pp_ne[grepl("[[:alnum:]]", pp_ne$EXPL)=="FALSE",]
        exploitant="agent"


        agents_flux_pp=unique(pp_ne$EXPL[grepl("[[:alnum:]]", pp_ne$EXPL)=="TRUE"])
        agents_flux_pm=unique(pm_ne$EXPL[grepl("[[:alnum:]]", pm_ne$EXPL)=="TRUE"])


        etat_expl=data.frame(Agents="",`Nbre clients concernés`="",`Lieu de Naissance`="", Profession ="", Codape="", Revenu = "", NIN="", TEL="",Adresse="",DATNAIS="", DATVALID="", ORIGINE_REV="", PAYS_RESID="", `Taux de fiabilisation`="", Appréciation="") # ##, `Taux à rattrapper`=taux_ratt)
        taux_filiale = lapply(agents, taux_function_pp )
        
        taux_filiale_t=bind_results(taux_filiale, etat_expl)
        taux_filiale_t = drop_lignes_vides(taux_filiale_t, "NIN")
        

        taux_completude_pp = taux_completude_table(taux_filiale_t, premier_jour, "F", "P")




   



      # minimum = taux_filiale_t[,c(2:8)]
      # minimum_1=c()
      # for (i in 1:ncol (minimum)) {
      #     minimume = gsub("%","",minimum[,i])
      #     minimum_1=cbind(minimum_1, as.numeric(minimume))
      # }
      # minimum_2 = c(paste("BOA", fil),min(minimum_1[,1]),min(minimum_1[,2]),min(minimum_1[,3]),min(minimum_1[,4]),min(minimum_1[,5]),min(minimum_1[,6]),min(minimum_1[,7]))
      # minimum_pp = rbind(minimum_2,minimum_pp)



                  
           ## PM
        exploitant="agent"
        agents_pm=unique(pm_ne$EXPL[grepl("[[:alnum:]]", pm_ne$EXPL)=="TRUE"])
        agents_null_pm=pm_ne[grepl("[[:alnum:]]", pm_ne$EXPL)=="FALSE",]

        etat_expl=data.frame(Agents="",`Nbre clients concernés`="", AGEC="", Capital="", CA = "", Resultat = "", RCSNO = "",CODAPE= "", ORIGINE_REV="", TEL="", `Taux de fiabilisation`="", Appréciation="" )##`Taux à rattrapper`="")
        taux_filiale_pm = lapply(agents_pm,taux_function_pm)
        
        taux_filiale_pm=bind_results(taux_filiale_pm, etat_expl)
        taux_filiale_pm = drop_lignes_vides(taux_filiale_pm, "RCSNO")


        taux_completude_pm = taux_completude_table(taux_filiale_pm, premier_jour, "F", "M")



  
     
      cat("###############################################\n")
      cat("######### Fin traitement flux  \n")
      cat("##############################################\n")
        
     
        rm(pp_ne,pm_ne)


        ### LE stock

          cat("######### Chargement des données Stock de\n",sigle, "\n")
        # Importatations des donnees
     
          # pp_stock/pm_stock déjà en mémoire, pas de relecture des CSV _STOCK
          pp_ne <- pp_stock
            pp_ne$AGENCE=as.numeric(pp_ne$AGENCE)

          #
          pp_ne=select_pp_fields(pp_ne, pp_fields)
          colnames(pp_ne)=toupper(colnames(pp_ne))


          pm_ne <- pm_stock
            pm_ne$AGENCE=as.numeric(pm_ne$AGENCE)

          #
          pm_ne=select_pm_fields(pm_ne, pm_fields)
          colnames(pm_ne)=toupper(colnames(pm_ne))
  
          pp_ne$EXPL[is_alphanumeric(pp_ne$EXPL)=="FALSE"]=paste("DIR_AGENCE_",pp_ne$AGENCE[is_alphanumeric(pp_ne$EXPL)=="FALSE"],sep="")
          pm_ne$EXPL[is_alphanumeric(pm_ne$EXPL)=="FALSE"]=paste("DIR_AGENCE_",pm_ne$AGENCE[is_alphanumeric(pm_ne$EXPL)=="FALSE"],sep="")




          if (exists("pp_ne")=="TRUE" & exists("pm_ne")=="TRUE") {

              cat("######### Les données Stock de\n",sigle, "ont été chargées avec succès \n")

          } else {

              cat("######### Erreur: les données Stock de\n",sigle, "n'ont pas été bien chargées (Voir les formats) #####\n")

          }

        pp_ne_s=pp_ne
        pm_ne_s=pm_ne


      cat("###############################################\n")
      cat("######### Début traitement stock  \n")
      cat("##############################################\n")
        
        
        ##-----Appréciation du Stock-####
        
        if (trimestre_actuel=="1"){
          trimestre=30
          Faible=5
          Moyen=30
          
        } else if (trimestre_actuel=="2") {
          trimestre=60
          Faible=30
          Moyen=60
        }  else if (trimestre_actuel=="3") {
          trimestre=90
          Faible=60
          Moyen=90
        }  else {
          trimestre=95
          Faible=65
          Moyen=95
        }
              

        ## Statistique a l'échelle filiale (PP)
        a=field_completion_rate(pp_ne, "PAYNAIS", inc)
        b=field_completion_rate(pp_ne, "PROFESSION", inc)
        c=field_completion_rate(pp_ne, "SALAIRE", inc)
        d=field_completion_rate(pp_ne, "CODAPE", inc)
        e=field_completion_rate(pp_ne, "TEL", inc)
        f=field_completion_rate(pp_ne, "ADRESSE", inc)
        g=field_completion_rate(pp_ne, "NUMID", inc)
        h=field_completion_rate(pp_ne, "DATNAIS", inc)
        i=field_completion_rate(pp_ne, "DATVALID", inc)
        j=field_completion_rate(pp_ne, "ORIGINE_REVENU", inc)
        k=field_completion_rate(pp_ne, "PAYS_RESID", inc)

        T=length(c(a,b,c,d,e,f,g,h,i,j))
        
       
                      
      #  taux=floor(100*(1-sum(nrow(pp_ne[pp_ne$PAYNAIS=="" | is.na(pp_ne$PAYNAIS)=="TRUE",]),nrow(pp_ne[pp_ne$PROFESSION=="" | is.na(pp_ne$PROFESSION)=="TRUE",]),
      #           nrow(pp_ne[pp_ne$SALAIRE=="" | is.na(pp_ne$SALAIRE)=="TRUE",]), nrow(pp_ne[pp_ne$NUMID=="" | is.na(pp_ne$NUMID)=="TRUE",]),
      #           nrow(pp_ne[pp_ne$CODAPE=="" | is.na(pp_ne$CODAPE)=="TRUE",]),nrow(pp_ne[pp_ne$TEL=="" | is.na(pp_ne$TEL)=="TRUE",]),          
      #           nrow(pp_ne[pp_ne$DATNAIS=="" | is.na(pp_ne$DATNAIS)=="TRUE",]),
      #           nrow(pp_ne[pp_ne$ADRESSE=="" |is.na(pp_ne$ADRESSE)=="TRUE",]),   nrow(pp_ne[pp_ne$NUMID=="" |is.na(pp_ne$NUMID)=="TRUE",]), nrow(pp_ne[pp_ne$DATVALID=="" |is.na(pp_ne$DATVALID)=="TRUE" ,]), 
      #                     nrow(pp_ne[pp_ne$ORIGINE_REV=="" |is.na(pp_ne$ORIGINE_REV)=="TRUE" ,])
      #           
      #           )/(T*nrow(pp_ne))))
      # 
      
       
        champs=c("PAYNAIS","PROFESSION","SALAIRE","NUMID","CODAPE","TEL","DATNAIS","ADRESSE","DATVALID","ORIGINE_REVENU","PAYS_RESID")

        taux=compute_taux(pp_ne, champs)
        appreciation=get_appreciation(taux, Faible, Moyen)
        taux=format_percent(taux)

        lieu_naiss=format_percent(a)
        profession=format_percent(b)
        revenu=format_percent(c)
        codape=format_percent(d)
        tel=format_percent(e)
        adresse=format_percent(f)
        nin=format_percent(g)
        datnais=format_percent(h)
        datvalid=format_percent(i)
        origine=format_percent(j)
        pays_resid=format_percent(k)



       
        etat_pp_nes=data.frame(`Lieu de Naissance` = lieu_naiss, Profession = profession,Codape=codape, Revenu = revenu, NIN = nin, TEL=tel, Adresse=adresse,DATNAIS=datnais,DATVALID=datvalid, ORIGINE_REV=origine, PAYS_RESID=pays_resid, `Taux de fiabilisation` =taux, Appréciation=appreciation)##`Taux à rattrapper` = taux_ratt)
        pp=etat_pp_nes[,c(ncol(etat_pp_nes),ncol(etat_pp_nes)-1,ncol(etat_pp_nes)-2)]
        pp_t=data.frame(A="",V="")

        colnames(pp_t)=c(paste(fil),"")
        pp_t[1,]=c("PM","PP")
        pp_t[2,2]=etat_pp_nes$Taux.de.fiabilisation[1]
        pp_t[3,2]=etat_pp_nes$Appréciation[1]

        rownames(pp_t)=c("","Taux de fiabilisation","Appréciation")

        

        ## Statistique a l'échelle filiale (PM)
            a=field_completion_rate(pm_ne, "CAPITAL", inc)
            b=field_completion_rate(pm_ne, "CA", inc)
            c=field_completion_rate(pm_ne, "RESULTAT", inc)
            d=field_completion_rate(pm_ne, "RCSNO", inc)
            e=field_completion_rate(pm_ne, "CODAPE", inc)
            f=field_completion_rate(pm_ne, "AGEC", inc)
            g=field_completion_rate(pm_ne, "ORIGINE_REVENU", inc)
            h=field_completion_rate(pm_ne, "TEL", inc)

           
            K=length(c(a,b,c,d,e))

           #taux=floor(100*(1-sum(nrow(pm_ne[pm_ne$AGEC %in% inc | is.na(pm_ne$AGEC)=="TRUE",]), nrow(pm_ne[pm_ne$CAPITAL %in% inc | is.na(pm_ne$CAPITAL)=="TRUE",]), nrow(pm_ne[pm_ne$CA %in% inc | is.na(pm_ne$CA)=="TRUE",]),
           #  nrow(pm_ne[pm_ne$RESULTAT %in% inc | is.na(pm_ne$RESULTAT)=="TRUE",]), nrow(pm_ne[pm_ne$RCSNO %in% inc | is.na(pm_ne$RCSNO=="TRUE"),]),
           #  nrow(pm_ne[pm_ne$CODAPE %in% inc | is.na(pm_ne$CODAPE=="TRUE"),]), nrow(pm_ne[pm_ne$ORIGINE_REV %in% inc | is.na(pm_ne$ORIGINE_REV)=="TRUE",]), nrow(pm_ne[pm_ne$TEL %in% inc | is.na(pm_ne$TEL)=="TRUE",]))/(K*nrow(pm_ne))))

                   champs=c("CAPITAL","CA","RESULTAT","RCSNO","CODAPE","AGEC","ORIGINE_REVENU","TEL")

        taux=compute_taux(pm_ne, champs)
        appreciation=get_appreciation(taux, Faible, Moyen)
        taux=format_percent(taux)
        capital=format_percent(a)
        ca=format_percent(b)
        resultat=format_percent(c)
        rcsno=format_percent(d)
        codape=format_percent(e)
        agec=format_percent(f)
        origine=format_percent(g)
        tel=format_percent(h)
            
            etat_pm_nes=data.frame(AGEC=agec,`Capital`=capital, CA = ca, Resultat = resultat, RCSNO = rcsno,CODAPE= codape,  ORIGINE_REV=origine, TEL=tel, `Taux de fiabilisation`=taux, Appréciation=appreciation)##`Taux à rattrapper`=taux_ratt)
            
        pm=etat_pm_nes[,c(ncol(etat_pm_nes),ncol(etat_pm_nes)-1,ncol(etat_pm_nes)-2)]
        
       

        pp_t[2,1]=c(pm[1,c(2)])
        pp_t[3,1]=c(pm[1,c(1)])
     
        tableau_suivi_stock = cbind(tableau_suivi_stock,pp_t)
        tableau_suivi_stock = tableau_suivi_stock[,-1]    

        ## PP

        stock="y"
        exploitant="agent"
        agents=unique(pp_ne$EXPL[grepl("[[:alnum:]]", pp_ne$EXPL)=="TRUE"])
        agents_null_pp=pp_ne[grepl("[[:alnum:]]", pp_ne$EXPL)=="FALSE",]
        
        agents_stock_pp=unique(pp_ne$EXPL[grepl("[[:alnum:]]", pp_ne$EXPL)=="TRUE"])
        agents_stock_pm=unique(pm_ne$EXPL[grepl("[[:alnum:]]", pm_ne$EXPL)=="TRUE"])

        
        etat_expl=data.frame(Agents="",`Nbre clients concernés`="",`Lieu de Naissance`="", Profession ="", Codape="", Revenu = "", NIN="", TEL="",Adresse="",DATNAIS="", DATVALID="", ORIGINE_REV="", PAYS_RESID="", `Taux de fiabilisation`="", Appréciation="") # ##, `Taux à rattrapper`=taux_ratt)
        taux_filiale = lapply(agents,taux_function_pp)

        taux_filiale_t_stock=bind_results(taux_filiale, etat_expl)
        taux_filiale_t_stock = drop_lignes_vides(taux_filiale_t_stock, "NIN")

        taux_completude_stock_pp = taux_completude_table(taux_filiale_t_stock, premier_jour, "S", "P")



                ## PM

        agents_pm=unique(pm_ne$EXPL[grepl("[[:alnum:]]", pm_ne$EXPL)=="TRUE"])
        agents_null_pm=pm_ne[grepl("[[:alnum:]]", pm_ne$EXPL)=="FALSE",]

        etat_expl=data.frame(Agents="", `Nbre clients concernés`="", AGEC="",Capital="", CA = "", Resultat = "", RCSNO = "",CODAPE= "", ORIGINE_REV="", TEL=tel, `Taux de fiabilisation`="", Appréciation="")##`Taux à rattrapper`="")
        taux_filiale_pm_stock = lapply(agents_pm,taux_function_pm)
        
        taux_filiale_pm_stock=bind_results(taux_filiale_pm_stock, etat_expl)
        taux_filiale_pm_stock = drop_lignes_vides(taux_filiale_pm_stock, "RCSNO")
      

      
        taux_completude_stock_pm = taux_completude_table(taux_filiale_pm_stock, premier_jour, "S", "M")





         taux_completude = rbind(taux_completude_pp,taux_completude_pm,taux_completude_stock_pp,taux_completude_stock_pm)

         write.csv2(taux_completude, paste(sep="",chemin,fil,"taux_",fil,".csv"), row.names=F)



        ### Creation du fichier Excel de suivi
        wb=createWorkbook()

  
     

 ### Resume filiale
 

k=ncol(tableau_suivi)



   ### Taux fiabilisation filiale

         wb_filiale=createWorkbook()
         addWorksheet(wb_filiale,"Récapitulatif")


   ### Taux fiabilisation filiale

         wb_filiale=createWorkbook()
         addWorksheet(wb_filiale,"Récapitulatif")

           ### recap stock          
             writeData(wb_filiale,"Récapitulatif", x=tableau_suivi_t, startRow=5, startCol=3)
             addStyle(wb_filiale,"Récapitulatif",headerStyle,cols=3:5, rows=5)
               ### recap flux          
             writeData(wb_filiale,"Récapitulatif", x=tableau_suivi_stock_t, startRow=5, startCol=3)
             addStyle(wb_filiale,"Récapitulatif",headerStyle,cols=7:9, rows=5)
             
           
    


        ### Sheet Flux PP
        addWorksheet(wb,"Flux PP")

        writeData(wb,"Flux PP", x=taux_filiale_t, startRow=1, startCol=1)

        k=ncol(taux_filiale_t)

        headerStyle <- createStyle(
        fontSize = 14, fontColour = "white", halign = "left",
        fgFill = "#09982E", border = "TopBottom", borderColour = "black"
        )

        addStyle(wb,"Flux PP",headerStyle,cols=1:k, rows=1)


        bonStyle <- createStyle(halign = "right",
        fgFill = "#298904"
        )
        moyenStyle <- createStyle(halign = "right",
        fgFill = "#F8DF19"
        )
        lowStyle <- createStyle(halign = "right",
        fgFill = "#D90A0A"
        )
        apply_status_style(wb, "Flux PP", taux_filiale_t, value_col = 15)


      ### Recap Sheet Flux PP
        addWorksheet(wb_recap,"Agent Flux PP")


        writeData(wb_recap,"Agent Flux PP", x=taux_filiale_t, startRow=1, startCol=1)

        k=ncol(taux_filiale_t)

       
        addStyle(wb_recap,"Agent Flux PP",headerStyle,cols=1:k, rows=1)


        apply_status_style(wb_recap, "Agent Flux PP", taux_filiale_t, value_col = 15)



         ### Sheet Flux 
         
              if (is.null(taux_filiale_pm)=="FALSE") {
                    addWorksheet(wb,"Flux PM")

                    writeData(wb,"Flux PM", x=taux_filiale_pm, startRow=1, startCol=1)

                    k=ncol(taux_filiale_pm)

                    headerStyle <- createStyle(
                    fontSize = 14, fontColour = "white", halign = "left",
                    fgFill = "#09982E", border = "TopBottom", borderColour = "black"
                    )

                    addStyle(wb,"Flux PM",headerStyle,cols=1:k, rows=1)


                    apply_status_style(wb, "Flux PM", taux_filiale_pm, value_col = 12, col_good = 9, col_mid = 9, col_low = 12)

              }

         ### Recap Sheet Flux PM

      if (is.null(taux_filiale_pm)=="FALSE") {

        addWorksheet(wb_recap,"Agent Flux PM")

        writeData(wb_recap,"Agent Flux PM", x=taux_filiale_pm, startRow=1, startCol=1)

        k=ncol(taux_filiale_pm)

        headerStyle <- createStyle(
        fontSize = 14, fontColour = "white", halign = "left",
        fgFill = "#09982E", border = "TopBottom", borderColour = "black"
        )

        addStyle(wb_recap,"Agent Flux PM",headerStyle,cols=1:k, rows=1)


        apply_status_style(wb_recap, "Agent Flux PM", taux_filiale_pm, value_col = 12)

              }

        ### Sheet Stock PP
        addWorksheet(wb,"Stock PP")


        writeData(wb,"Stock PP", x=taux_filiale_t_stock, startRow=1, startCol=1)

        k=ncol(taux_filiale_t_stock)

        headerStyle <- createStyle(
        fontSize = 14, fontColour = "white", halign = "left",
        fgFill = "#09982E", border = "TopBottom", borderColour = "black"
        )

        addStyle(wb,"Stock PP",headerStyle,cols=1:k, rows=1)


        bonStyle <- createStyle(halign = "right",
        fgFill = "#298904"
        )
        moyenStyle <- createStyle(halign = "right",
        fgFill = "#F8DF19"
        )
        lowStyle <- createStyle(halign = "right",
        fgFill = "#D90A0A"
        )

        apply_status_style(wb, "Stock PP", taux_filiale_t_stock, value_col = 15)

          ### Recap Sheet Stock PP
        addWorksheet(wb_recap,"Agent Stock PP")


        writeData(wb_recap,"Agent Stock PP", x=taux_filiale_t_stock, startRow=1, startCol=1)

        k=ncol(taux_filiale_t_stock)

        headerStyle <- createStyle(
        fontSize = 14, fontColour = "white", halign = "left",
        fgFill = "#09982E", border = "TopBottom", borderColour = "black"
        )

        addStyle(wb_recap,"Agent Stock PP",headerStyle,cols=1:k, rows=1)


        bonStyle <- createStyle(halign = "right",
        fgFill = "#298904"
        )
        moyenStyle <- createStyle(halign = "right",
        fgFill = "#F8DF19"
        )
        lowStyle <- createStyle(halign = "right",
        fgFill = "#D90A0A"
        )

        apply_status_style(wb_recap, "Agent Stock PP", taux_filiale_t_stock, value_col = 15)

         ### Sheet Stock PM
        addWorksheet(wb,"Stock PM")

        writeData(wb,"Stock PM", x=taux_filiale_pm_stock, startRow=1, startCol=1)

        k=ncol(taux_filiale_pm_stock)

        headerStyle <- createStyle(
        fontSize = 14, fontColour = "white", halign = "left",
        fgFill = "#09982E", border = "TopBottom", borderColour = "black"
        )

        addStyle(wb,"Stock PM",headerStyle,cols=1:k, rows=1)


        apply_status_style(wb, "Stock PM", taux_filiale_pm_stock, value_col = 12)


         ### Recap Stock PM
        addWorksheet(wb_recap," Agent Stock PM")

        writeData(wb_recap," Agent Stock PM", x=taux_filiale_pm_stock, startRow=1, startCol=1)

        k=ncol(taux_filiale_pm_stock)

        headerStyle <- createStyle(
        fontSize = 14, fontColour = "white", halign = "left",
        fgFill = "#09982E", border = "TopBottom", borderColour = "black"
        )

        addStyle(wb_recap," Agent Stock PM",headerStyle,cols=1:k, rows=1)


        apply_status_style(wb_recap, " Agent Stock PM", taux_filiale_pm_stock, value_col = 9)


            saveWorkbook(wb, paste(sep="",paste(chemin,fil,"//",sep=""),"Rapport des taux de complétude par agent de ",sigle,".xlsx"), overwrite=T)
            
              saveWorkbook(wb_filiale, paste(sep="",paste(chemin,fil,"//",sep=""),"Rapport du taux de complétude de la filiale ",sigle,".xlsx"), overwrite=T)
          
              saveWorkbook(wb_recap, paste(sep="",paste(chemin,fil,"//",sep=""),"Taux de complétude BOA_",sigle,".xlsx"), overwrite=T)


         
            
            if (exists("tableau_fiabilisation")=="TRUE") {
              cat("######################################################################################\n")
              cat("######### Les fichiers des suivi de ",sigle," par agence sont générés avec sucès ######\n")
              cat("######################################################################################\n")
            } else {
              cat("######################################################################################\n")
              cat("######### ERREUR: Les fichiers des suivi de ",sigle," par agence ne sont pas générés ####\n")
              cat("######################################################################################\n")
              
            }


            


            ## les graphiques Fiabilisation PP

          # Ligne "Taux de fiabilisation" (2e ligne) des 2 dernieres colonnes (PM, PP).
          # Robuste que le tableau ait ou non sa colonne de libelles en tete.
          taux_row <- function(tab) {
            tab <- as.data.frame(tab)
            n <- ncol(tab)
            if (n < 2 || nrow(tab) < 2) return(c(NA_character_, NA_character_))
            as.character(unlist(tab[2, (n-1):n, drop = FALSE], use.names = FALSE))
          }

          v_flux  <- taux_row(tableau_suivi_t)
          v_stock <- taux_row(tableau_suivi_stock_t)

          suivi_fiabilisation_i <- data.frame(
            "Flux PM"  = v_flux[1],
            "Flux PP"  = v_flux[2],
            "Stock PM" = v_stock[1],
            "Stock PP" = v_stock[2],
            Date       = as.Date(jour_recent) - 1,
            check.names = FALSE, stringsAsFactors = FALSE
          )

         
          fichier_suivi <- paste0(chemin,fil,"//suivi_fiabilisation.txt")

          # L'historique ne doit JAMAIS etre perdu : s'il est illisible ou n'a pas
          # le format attendu, on n'ecrase rien et on signale.
          historique_ok <- !is.null(suivi_fiabilisation) &&
                           is.data.frame(suivi_fiabilisation) &&
                           ncol(suivi_fiabilisation) == 5

          if (!historique_ok && !is.null(suivi_fiabilisation) &&
              is.data.frame(suivi_fiabilisation) && ncol(suivi_fiabilisation) > 0) {
            cat("######### ATTENTION", fil, ": suivi_fiabilisation.txt a",
                ncol(suivi_fiabilisation), "colonnes au lieu de 5 -> fichier NON modifie\n")
          } else {
            if (historique_ok) {
              colnames(suivi_fiabilisation)=c("Flux PM","Flux PP","Stock PM", "Stock PP","Date")
              suivi_fiabilisation$Date=parse_date_any(suivi_fiabilisation$Date)
            } else {
              # Pas d'historique (premier passage) : on part de la ligne du mois
              suivi_fiabilisation <- suivi_fiabilisation_i[0, , drop = FALSE]
            }

            suivi_fiabilisation=rbind(suivi_fiabilisation,suivi_fiabilisation_i)

            # Sauvegarde de l'historique avant reecriture
            if (file.exists(fichier_suivi)) {
              file.copy(fichier_suivi,
                        paste0(chemin,fil,"//suivi_fiabilisation_backup.txt"),
                        overwrite = TRUE)
            }

            suivi_fiabilisation_out <- suivi_fiabilisation
            suivi_fiabilisation_out$Date <- format(suivi_fiabilisation_out$Date, "%d/%m/%Y")
            suivi_fiabilisation_out =unique(suivi_fiabilisation_out)

            write.csv2(suivi_fiabilisation_out, fichier_suivi)

            write.csv2(suivi_fiabilisation_out, paste0(chemin,fil,"//data//suivi_fiabilisation_",fil,".csv"), row.names=FALSE)
          }


          





     ## Les taux de complétude par agent flux et stock
     # Calcul direct depuis les tableaux déjà en mémoire (taux_completude_*),
     # sans relire le classeur "Rapport des taux de complétude par agent" (lent).

          write.csv2(rbind(taux_completude_pp, taux_completude_pm, taux_completude_stock_pp, taux_completude_stock_pm),
                     paste0(chemin,fil,"//data//taux_",fil,".csv"), row.names=FALSE)


          

     data_dir <- paste0(chemin, fil, "//data//")

dossier_A <- paste0(chemin,fil,"//data//")
dossier_B <- chemin_python


fichiers <- c(
  paste0("taux_", fil, ".csv"),
  paste0("suivi_fiabilisation_", fil, ".csv"),
  paste0("agents_", fil, ".csv"),
  paste0("pp_", fil, "_STOCK_F.csv"),
  paste0("pm_", fil, "_STOCK_F.csv"),
  paste0("pp_", fil, "_STOCK.csv"),
  paste0("pm_", fil, "_STOCK.csv")
)

# 3. Construire les chemins complets
chemins_sources <- file.path(dossier_A, fichiers)
chemins_destinations <- file.path(dossier_B, fichiers)

# 4. Copier les fichiers
# Utilisation de 'overwrite = TRUE' si vous voulez remplacer les fichiers existants dans B


file.copy(from = chemins_sources, to = chemins_destinations, overwrite = TRUE)
  cat("######################################################################################\n")
              cat("######### Génération du rapport de suivi de ",sigle," pour la Direction générale ######\n")
              cat("######################################################################################\n")



 sigle <- paste0("BOA_", fil)
  ppt <- dg_pptx()                   # presentation vierge 16/9, aucun template externe
  W <- slide_size(ppt)$width
  H <- slide_size(ppt)$height
  M <- 0.3


    # 0. Configuration des libellés EXACTS de l'image

  # 1. Chargement et nettoyage des données (PP et PM)
  # pp_stock/pm_stock déjà en mémoire, pas de relecture des CSV _STOCK
  pp_ne <- clean_colnames(pp_stock)
  pp_ne$AGENCE <- as.numeric(pp_ne$AGENCE)
  colnames(pp_ne) <- toupper(colnames(pp_ne))

  pm_ne <- clean_colnames(pm_stock)
  pm_ne$AGENCE <- as.numeric(pm_ne$AGENCE)
   pm_ne=select_pm_fields(pm_ne, pm_fields)
  colnames(pm_ne) <- toupper(colnames(pm_ne))
  
  # Correction EXPL et Marquage Anomalies
  pp_ne$EXPL[is_alphanumeric(pp_ne$EXPL) == FALSE] <- paste0("DIR_AGENCE_", pp_ne$AGENCE[is_alphanumeric(pp_ne$EXPL) == FALSE])
  pm_ne$EXPL[is_alphanumeric(pm_ne$EXPL) == FALSE] <- paste0("DIR_AGENCE_", pm_ne$AGENCE[is_alphanumeric(pm_ne$EXPL) == FALSE])
  pp_ne <- clean_and_mark_anomalies(pp_ne, start_col = 4)
  pm_ne <- clean_and_mark_anomalies(pm_ne, start_col = 4)
  
  # Filtre Flux
  pp_flux <- pp_ne %>% filter(DATOUV >= as.Date(premier_jour), DATOUV < as.Date(jour_recent))
  pm_flux <- pm_ne %>% filter(DATOUV >= as.Date(premier_jour), DATOUV < as.Date(jour_recent))

  # pm_exclure a ete calcule a l'import (cf. section "3 bis"), a partir de la
  # colonne pm_exclure de prerequis.csv.
  pm_sans_rcsno_stock <- pm_without_rcsno(pm_ne, pm_exclure)
  pm_sans_rcsno_flux <- pm_without_rcsno(pm_flux, pm_exclure)

  readr::write_delim(pm_sans_rcsno_stock,
                     paste0(chemin, fil, "//data//pm_", fil, "_sans_RCSNO_STOCK.csv"),
                     delim = ";", quote = "all", escape = "double", na = "")
  readr::write_delim(pm_sans_rcsno_flux,
                     paste0(chemin, fil, "//data//pm_", fil, "_sans_RCSNO_FLUX.csv"),
                     delim = ";", quote = "all", escape = "double", na = "")

  pp_datnais_01011900_stock <- pp_birthdate_01011900(pp_ne)
  readr::write_delim(pp_datnais_01011900_stock,
                     paste0(chemin, fil, "//data//pp_", fil, "_DATNAIS_01011900_STOCK.csv"),
                     delim = ";", quote = "all", escape = "double", na = "")
  

  # 2 bis. Indicateurs calcules UNE SEULE FOIS
  # Ces taux sont reutilises par les tableaux des slides, les KPI et les
  # observations. Chaque appel balaie l'integralite du stock : on les evalue
  # ici une fois pour toutes au lieu de les recalculer a chaque usage.
  stock_pp <- setNames(sapply(pp_fields, function(f) champ_rate(pp_ne, f)),   pp_labels)
  stock_pm <- setNames(sapply(pm_fields, function(f) champ_rate(pm_ne, f)),   pm_labels)
  flux_pp  <- setNames(sapply(pp_fields, function(f) champ_rate(pp_flux, f)), pp_labels)
  flux_pm  <- setNames(sapply(pm_fields, function(f) champ_rate(pm_flux, f)), pm_labels)

  g_pp      <- global_rate(pp_ne, pp_fields)
  g_pm      <- global_rate(pm_ne, pm_fields)
  g_pp_flux <- global_rate(pp_flux, pp_fields)
  g_pm_flux <- global_rate(pm_flux, pm_fields)

  rates_pp <- agent_completion_rates(pp_ne, pp_fields)
  rates_pm <- agent_completion_rates(pm_ne, pm_fields)

  # 3. Construction des Dataframes pour PP
  pp_table_ppt <- data.frame(
    Col1 = c(rep("Taux par champs", length(pp_labels)), rep("Taux de complétude global", 2)),
    Champs = c(pp_labels, "Approche par champs (1)", "Approche par clients (2)"),
    Flux = c(unname(sapply(flux_pp, to_pct)),
             to_pct(g_pp_flux),
             to_pct(client_completion_rate(pp_flux, pp_fields))),
    Stock = c(unname(sapply(stock_pp, to_pct)),
              to_pct(g_pp),
              to_pct(client_completion_rate(pp_ne, pp_fields))),
    stringsAsFactors = FALSE
  )
  colnames(pp_table_ppt) <- c(" ", "Champs", "Flux (3)", "Stock (4)")

  # 4. Construction des Dataframes pour PM
  pm_table_ppt <- data.frame(
    Col1 = c(rep("Taux par champs", length(pm_labels)), rep("Taux de complétude global", 2)),
    Champs = c(pm_labels, "Approche par champs (1)", "Approche par clients (2)"),
    Flux = c(unname(sapply(flux_pm, to_pct)),
             to_pct(g_pm_flux),
             to_pct(client_completion_rate(pm_flux, pm_fields))),
    Stock = c(unname(sapply(stock_pm, to_pct)),
              to_pct(g_pm),
              to_pct(client_completion_rate(pm_ne, pm_fields))),
    stringsAsFactors = FALSE
  )
  colnames(pm_table_ppt) <- c(" ", "Champs", "Flux (3)", "Stock (4)")

  # 5. Application du style
  ft_pp <- style_ft_exact(pp_table_ppt)
  ft_pm <- style_ft_exact(pm_table_ppt)

          notation_fil=notation[notation$Filiale==paste0("BOA ",fil),]
          if (nrow(notation_fil) == 0) {
            notation_fil <- data.frame(
              Filiale = paste0("BOA ", fil),
              Nombre_notes = 0,
              Insuffisant = "0%",
              Passable = "0%",
              Bien = "0%",
              "Très Bien" = "0%",
              check.names = FALSE
            )
          } else {
            notation_fil <- notation_fil[, c("Filiale", "Nombre_notes", note_cols), drop = FALSE]
          }
          colnames(notation_fil)=c("Filiale","Nombre notes","Insuffisant","Passable","Bien","Très Bien")

          repartition_pp <- agent_distribution_row(rates_pp)
          repartition_pm <- agent_distribution_row(rates_pm)
          taux_compl <- as.data.frame(rbind(repartition_pp, repartition_pm), check.names = FALSE)
          taux_compl <- data.frame(
            Taux = c("Taux stock PP", "Taux stock PM"),
            taux_compl,
            check.names = FALSE
          )

          ft_taux_compl <- style_slide5_table(taux_compl, widths = c(2.5, 1.35, 1.35, 1.35, 1.35))
          ft_notation <- style_slide5_table(notation_fil, widths = c(1.45, 1.35, 1.3, 1.3, 1.3, 1.3))

# Nombre de clients non fiabilisés

  # non_fiab est deja charge et filtre (CODE == "OI") a l'import : on le
  # reutilise tel quel plutot que de relire le CSV.


  #####--------- CALCUL DES OBSERVATIONS (avant-derniere slide) ---------#####
  # Toutes les observations sont chiffrees a partir des donnees de la filiale.

  # Libelle de periode : plage complete en mode "mois", date unique en mode
  # "jour" (cf. bornes_periode). Repli calcule si la variable n'existe pas.
  periode_lib <- if (exists("periode_libelle") && !is.null(periode_libelle)) {
    periode_libelle
  } else {
    paste0("du ", format(as.Date(premier_jour), "%d/%m/%Y"),
           " au ", format(as.Date(jour_recent) - 1, "%d/%m/%Y"))
  }
  date_arrete <- format(as.Date(jour_recent) - 1, "%d/%m/%Y")

  # rates_pp / rates_pm et les taux par champ (stock_*, flux_*) sont calcules
  # plus haut, section "2 bis".
  pct80 <- function(r) if (length(r) == 0) 0 else round(100 * mean(r > 80))
  nb80  <- function(r) sum(r > 80)

  # Champs hors seuil : stock < 90%, flux < 100%
  hs_stock_pp <- stock_pp[!is.na(stock_pp) & stock_pp < 90]
  hs_stock_pm <- stock_pm[!is.na(stock_pm) & stock_pm < 90]
  hs_flux_pp  <- flux_pp[!is.na(flux_pp)  & flux_pp  < 100]
  hs_flux_pm  <- flux_pm[!is.na(flux_pm)  & flux_pm  < 100]

  # Notation du Controle permanent
  n_notes  <- suppressWarnings(as.numeric(notation_fil[["Nombre notes"]]))[1]
  if (is.na(n_notes)) n_notes <- 0
  p_bien   <- sum(dg_num(notation_fil[["Bien"]]), dg_num(notation_fil[["Très Bien"]]), na.rm = TRUE)
  p_faible <- sum(dg_num(notation_fil[["Passable"]]), dg_num(notation_fil[["Insuffisant"]]), na.rm = TRUE)
  n_agents <- length(unique(c(trimws(as.character(pp_ne$EXPL)), trimws(as.character(pm_ne$EXPL)))))
  couv <- if (n_agents > 0) round(100 * n_notes / n_agents) else 0

  obs <- c(
    sprintf("Complétude globale au %s : PP %s en stock (%s en flux), PM %s en stock (%s en flux).",
            date_arrete,
            to_pct(g_pp), to_pct(g_pp_flux),
            to_pct(g_pm), to_pct(g_pm_flux)),

    sprintf("Chargés de compte au-dessus du seuil de 80 %% : %d %% en PP (%d agents) et %d %% en PM (%d agents), sur %d chargés de compte actifs.",
            pct80(rates_pp), nb80(rates_pp), pct80(rates_pm), nb80(rates_pm), n_agents),

    if (length(hs_stock_pp) > 0)
      sprintf("PP - champs sous le seuil stock de 90 %% : %s.", dg_liste_champs(hs_stock_pp))
    else "PP - tous les champs critiques atteignent le seuil stock de 90 %.",

    if (length(hs_stock_pm) > 0)
      sprintf("PM - champs sous le seuil stock de 90 %% : %s.", dg_liste_champs(hs_stock_pm))
    else "PM - tous les champs critiques atteignent le seuil stock de 90 %.",

    if (length(hs_flux_pp) + length(hs_flux_pm) > 0)
      sprintf("Flux non conforme au seuil de 100 %% : %s.",
              paste(na.omit(c(if (length(hs_flux_pp)) paste0("PP ", dg_liste_champs(hs_flux_pp)),
                              if (length(hs_flux_pm)) paste0("PM ", dg_liste_champs(hs_flux_pm)))),
                    collapse = " / "))
    else "Flux conforme : 100 % de complétude sur l'ensemble des champs critiques PP et PM.",

    if (n_notes > 0)
      sprintf("Contrôle permanent : %d agents notés (%d %% des chargés de compte) ; %g %% notés Bien ou Très Bien, %g %% à améliorer (Passable ou Insuffisant).",
              n_notes, couv, p_bien, p_faible)
    else "Contrôle permanent : aucun agent noté sur la plateforme KYC pour cette période.",

    sprintf("%d comptes sous interdit crédit et débit sont exclus du périmètre à fiabiliser.",
            nrow(non_fiab))
  )


  # Cartes KPI de la slide Observations.
  # Une couleur vive par carte ; bascule sur DG_ALERTE si le seuil n'est pas tenu.
  kpi_col <- function(i, v = NULL, seuil = NULL) {
    if (!is.null(seuil) && (is.na(v) || v < seuil)) DG_ALERTE else DG_FUN[i]
  }
  n_hs <- length(hs_stock_pp) + length(hs_stock_pm)

  kpis <- list(
    list(valeur = to_pct(g_pp), libelle = "COMPLÉTUDE PP — STOCK",
         couleur = kpi_col(1, g_pp, 90)),
    list(valeur = to_pct(g_pm), libelle = "COMPLÉTUDE PM — STOCK",
         couleur = kpi_col(2, g_pm, 90)),
    list(valeur = paste0(pct80(rates_pp), "%"), libelle = "CC AU-DESSUS DE 80 % — PP",
         couleur = kpi_col(3, pct80(rates_pp), 80)),
    list(valeur = paste0(pct80(rates_pm), "%"), libelle = "CC AU-DESSUS DE 80 % — PM",
         couleur = kpi_col(4, pct80(rates_pm), 80)),
    list(valeur = as.character(n_hs), libelle = "CHAMPS SOUS SEUIL",
         couleur = if (n_hs == 0) DG_FUN[5] else DG_ALERTE),
    list(valeur = as.character(n_notes), libelle = "AGENTS NOTÉS PAR LE CP",
         couleur = DG_FUN[6])
  )


  #####--------- REMPLISSAGE DE LA MAQUETTE ---------#####
  # La maquette "exemple_rapport_kyc.pptx" fournit deja en-tetes, pieds de page,
  # logo, sommaire et pages de garde/fin. On n'y depose que le contenu.
  # Zone utile d'une slide de contenu : y de 0,70 a 5,30 ; x de 0,30 a 9,70.

  note_1 <- "(1) Taux de complétude par champs : proportion d'informations renseignées rapportée au nombre total d'informations attendues sur les champs critiques."
  note_2 <- "(2) Taux de complétude par clients : proportion de clients présentant au moins une information manquante sur les champs critiques."
  note_3 <- "(3) Le taux de complétude doit être égal à 100 % pour le flux."
  note_4 <- "(4) Le taux de complétude doit être supérieur ou égal à 90 % pour le stock."
  notes_bas <- block_list(dg_par(note_1, size = 7, color = DG_MUET),
                          dg_par(note_2, size = 7, color = DG_MUET),
                          dg_par(note_3, size = 7, color = DG_MUET),
                          dg_par(note_4, size = 7, color = DG_MUET))

  # Ligne de periode, en haut a droite des slides de contenu
  bandeau_periode <- function(ppt, txt) {
    ph_with(ppt, value = block_list(
        dg_par(txt, size = 9, bold = TRUE, color = DG_MUET, align = "right")),
      location = ph_location(left = 4.9, top = 0.70, width = 4.8, height = 0.26))
  }

  ## Slide 1 - page de garde (fond vert) : titre et periode
  ppt <- on_slide(ppt, index = 1)
  ppt <- ph_with(ppt, value = block_list(
      dg_par(paste0("FIABILISATION KYC BOA ", fil), size = 30, bold = TRUE, color = "white"),
      dg_par("CONTRÔLE ET SUIVI DES TRAVAUX DE FIABILISATION KYC", size = 13, color = "#C8F0D8")),
    location = ph_location(left = 0.50, top = 1.95, width = 7.2, height = 1.15))
  ppt <- ph_with(ppt, value = block_list(
      dg_par(paste0("Données ", periode_lib), size = 12, bold = TRUE, color = "white")),
    location = ph_location(left = 0.50, top = 3.40, width = 7.0, height = 0.34))

  ## Slide 3 - completude PP
  ppt <- on_slide(ppt, index = 3)
  ppt <- bandeau_periode(ppt, paste0("Données ", periode_lib))
  ppt <- ph_with(ppt, value = ft_pp,
                 location = ph_location(left = 0.30, top = 1.00, width = 9.40, height = 3.38))
  ppt <- ph_with(ppt, value = notes_bas,
                 location = ph_location(left = 0.30, top = 4.46, width = 9.40, height = 0.84))

  ## Slide 4 - completude PM
  ppt <- on_slide(ppt, index = 4)
  ppt <- bandeau_periode(ppt, paste0("Données ", periode_lib))
  ppt <- ph_with(ppt, value = ft_pm,
                 location = ph_location(left = 0.30, top = 1.00, width = 9.40, height = 3.38))
  ppt <- ph_with(ppt, value = notes_bas,
                 location = ph_location(left = 0.30, top = 4.46, width = 9.40, height = 0.84))

  ## Slide 5 - repartition des taux + notation
  ppt <- on_slide(ppt, index = 5)
  ppt <- bandeau_periode(ppt, paste0("Données ", periode_lib))
  ppt <- ph_with(ppt, value = dg_bloc(list(
      list(txt = "1.  RÉPARTITION DES CHARGÉS DE COMPTE PAR TRANCHE DE TAUX (STOCK)",
           size = 10, bold = TRUE, color = "white", bg = DG_VERT, h = 0.30)), 9.40, pad_left = 10),
    location = ph_location(left = 0.30, top = 1.05, width = 9.40, height = 0.30))
  ppt <- ph_with(ppt, value = ft_taux_compl,
                 location = ph_location(left = 0.42, top = 1.48, width = 9.16, height = 1.05))
  ppt <- ph_with(ppt, value = dg_bloc(list(
      list(txt = "2.  RÉPARTITION DES NOTES ATTRIBUÉES PAR LE CONTRÔLE PERMANENT",
           size = 10, bold = TRUE, color = "white", bg = DG_NAVY, h = 0.30)), 9.40, pad_left = 10),
    location = ph_location(left = 0.30, top = 2.95, width = 9.40, height = 0.30))
  ppt <- ph_with(ppt, value = ft_notation,
                 location = ph_location(left = 0.42, top = 3.38, width = 9.16, height = 0.90))

  ## Slide 6 - OBSERVATIONS : cartes KPI colorees + points d'attention
  ppt <- on_slide(ppt, index = 6)
  ppt <- bandeau_periode(ppt, paste0("BOA ", fil, "  |  ", periode_lib))
  ppt <- dg_kpi_row(ppt, kpis, W, top = 1.02, M = 0.30, hauteur = 0.88)
  ppt <- ph_with(ppt, value = dg_bloc(list(
      list(txt = "POINTS D'ATTENTION", size = 9, bold = TRUE,
           color = "white", bg = DG_NAVY, h = 0.26)), 9.40, pad_left = 10),
    location = ph_location(left = 0.30, top = 2.10, width = 9.40, height = 0.26))
  ppt <- ph_with(ppt, value = dg_puces(obs, size = 9),
                 location = ph_location(left = 0.35, top = 2.46, width = 9.30, height = 2.80))

   print(ppt,target=paste(chemin,fil,"//Rapport fiabilisation KYC BOA ",fil," _ V2 ",format(Sys.Date(),"%B %Y"),".pptx",sep=""))

  cat("PPT généré avec succès pour", fil, "\n")

}



for (fil in filiale){
  tryCatch(
    production(fil),
    error = function(e) {
      message("Erreur filiale ", fil, " : ", e$message)
    }
  )
}