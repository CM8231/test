# -*- coding: utf-8 -*-
import sys

def generation_rapport(dernierIdx):
    # Fonction permettant de récupérer les résultats de la table matlab et de les intégrer dans un rapport qualité pdf
	#dernierIdx =130
        # Import des modules externes
    import scipy.io as sio
    import numpy as np
    from datetime import datetime
    import textwrap
    import argparse
    import os.path
    from sys import exit, path, version
    from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.lib.pagesizes import A4, landscape, portrait
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
    import math
    import glob
    from pickle import load, dump
    from os import mkdir, system, environ, remove
    from os.path import isfile, isdir
    from shutil import copyfile
    from collections import OrderedDict
    from IRMAGE_reporting import PageNumCanvas, ReportLine
	
    Path = '/media/cmaggia/Mesfichiers/MATLAB/'
    NomTraiteur = 'MAGGIA Christophe'
    # os.chdir(Path)
    test = sio.loadmat(Path + 'Basedonneesv4.mat')
    tableau = test['BaseDonnees']
    dernierIdx = dernierIdx-1
    # for dernierIdx  in range(len(tableau)):
    centre = tableau[dernierIdx]['Centre'][0][0].encode('utf-8')
    MarqueMachine = tableau[dernierIdx]['MarqueMachine'][0][0].encode('utf-8')
    ChampMachine = str(tableau[dernierIdx]['ChampMachine'][0][0][0])
    date_now = datetime.now()
    SeuilVol=0.15
    SeuilCoupe=0.1
    dGen = OrderedDict()
    dGen.update({'Code Sujet': tableau[dernierIdx]['NomSujet'][0][0].encode('utf-8')})
    sujet = dGen['Code Sujet']
    print(sujet)
    age = tableau[dernierIdx]['Age'][0][0][0]
    if age < 2 or age > 100:
        dGen.update({'Age': 'Inconnu'})
    else:
        dGen.update({'Age': str(age)})

    if tableau[dernierIdx]['Sexe'][0].size > 0 and 'M' in tableau[dernierIdx]['Sexe'][0] or 'F' in tableau[dernierIdx]['Sexe'][0]:
        dGen.update({'Sexe': tableau[dernierIdx]['Sexe'][0][0].encode('utf-8')})
    else:
        dGen.update({'Sexe':  'Non renseigné'})

    dGen.update({'Date de l"examen': str(tableau[dernierIdx]['DateExam'][0][0][0])})

    if 'V' in sujet:
        dGen.update({'Etiologie': 'Volontaire sain'})
    elif 'T' in sujet:
        dGen.update({'Etiologie': 'Patient TC test'})
    elif 'P' in sujet:
        dGen.update({'Etiologie': 'Patient TC'})

    Cheminsave = Path + centre + '/Sujets/' + sujet +'/'

    logo_OxyTC = Image(Path + 'Programmes/logoOxyTC.jpg', 30.0*mm, 30.0*mm)
    logo_CHU = Image(Path + 'Programmes/logoCHU.png', 25.0*mm, 25.0*mm)
    logo_Gin = Image(Path + 'Programmes/logo-GIN.png', 50.0*mm, 25.0*mm)
    logo_Irmage = Image(Path + 'Programmes/logoIrmage.jpg', 25.0*mm, 25.0*mm)
    logo_OK = Image(Path + 'Programmes/ok-button.png', 8.0*mm, 8.0*mm)
    logo_refus = Image(Path + 'Programmes/Refus.jpg', 7.0*mm, 7.0*mm)

    page = SimpleDocTemplate(Cheminsave + sujet + '.pdf', pagesize=portrait(A4), rightMargin=20*mm, leftMargin=20*mm, topMargin=10*mm, bottomMargin=10*mm) # Output document template definition for page; margins and pagessize.

    styles = getSampleStyleSheet()					# Initialises stylesheet with few basic heading and text styles, return a stylesheet object.
    styles.add(ParagraphStyle(name='Center', alignment=TA_CENTER))	# ParagraphStyle gives all the attributes available for formatting paragraphs.
    styles.add(ParagraphStyle(name='CenterRed', alignment=TA_CENTER, textColor=colors.red))
    styles.add(ParagraphStyle(name='Center2', alignment=TA_CENTER))
    styles['Center2'].leading = 24 # If more than 1 line.
    styles.add(ParagraphStyle(name='Justify', alignment=TA_JUSTIFY))
    styles.add(ParagraphStyle(name='Right', alignment=TA_RIGHT))
    styles.add(ParagraphStyle(name='Left', alignment=TA_LEFT))
    styles.add(ParagraphStyle(name='Bullet1', leftIndent=30, bulletOffsetY=2, bulletIndent=20, bulletFontSize=6, bulletColor='black', bulletText=u'●'))
    styles.add(ParagraphStyle(name='Bullet2', leftIndent=60, bulletOffsetY=1, bulletIndent=50, bulletFontSize=6, bulletColor='black', bulletText=u'❍'))
    styles.add(ParagraphStyle(name='Indent', leftIndent=30))
    styles.add(ParagraphStyle(name='Indent2', leftIndent=60))
    styles.add(ParagraphStyle(name='Indent2Red', leftIndent=60, textColor=colors.red))
    rapport = []
    logo_OxyTC.hAlign = 'LEFT'
    logo_CHU.hAlign = 'LEFT'
    logo_Gin.hAlign = 'LEFT'
    logo_Irmage.hAlign = 'LEFT'

    listeSeq = ['T1', 'FLAIR', 'T2ETOILE', 'T2HR', 'B0', 'DTI']

    dSeq = OrderedDict()
    didSeq = OrderedDict()
    #dlisteNoire = OrderedDict()
    #dCQstatut = OrderedDict()
    for seq in listeSeq:
        dSeq.update({seq: []})
        didSeq.update({seq: []})
        #dCQstatut.update({seq: []})
        #dlisteNoire.update({seq: []})
        [didSeq.update({seq: s}) for s in range(len(tableau[dernierIdx]['Fichier'][0]['NomSeq'][0])) if len(tableau[dernierIdx]['Fichier'][0]['NomSeq'][0][s][0]) > 0 and seq == tableau[dernierIdx]['Fichier'][0]['NomSeq'][0][s][0].encode('utf-8')]
        [dSeq.update({seq: tableau[dernierIdx]['Fichier'][0]['Chemin'][0][s][0].encode('utf-8')}) for s in range(len(tableau[dernierIdx]['Fichier'][0]['NomSeq'][0])) if len(tableau[dernierIdx]['Fichier'][0]['NomSeq'][0][s][0]) > 0 and seq == tableau[dernierIdx]['Fichier'][0]['NomSeq'][0][s][0].encode('utf-8')]
        #[dCQstatut.update({seq: tableau[dernierIdx]['CQ'][0]['Statut'][0][s][0].encode('utf-8')}) for s in range(len(tableau[dernierIdx]['CQ'][0]['NomSeq'][0])) if len(tableau[dernierIdx]['CQ'][0]['NomSeq'][0][s]) > 0 and seq == tableau[dernierIdx]['CQ'][0]['NomSeq'][0][s][0].encode('utf-8')]
        #[dlisteNoire.update({seq: tableau[dernierIdx]['CQ'][0]['Anomalie'][0][s]}) for s in range(len(tableau[dernierIdx]['CQ'][0]['NomSeq'][0])) if len(tableau[dernierIdx]['CQ'][0]['NomSeq'][0][s]) > 0 and seq == tableau[dernierIdx]['CQ'][0]['NomSeq'][0][s][0].encode('utf-8')] and tableau[dernierIdx]['CQ'][0]['Statut'][0][s][0].encode('utf-8') == 'Invalide']

    for seq in dSeq.iterkeys():
        if len(dSeq[seq]) > 0:
            Chemin = dSeq[seq].split('/')
            Chemin.pop()
            Chemin = "/".join(Chemin)
            dSeq.update({seq: Chemin})
            del Chemin
    ############################################ En-tete ####################################################

    entete_logo = [[logo_OxyTC, logo_CHU, logo_Gin, logo_Irmage]]
    entete = Table(entete_logo)
    entete.hAlign = 'CENTER'
    rapport.append(entete)
    rapport.append(Spacer(0*mm, 2*mm))  # (width, height)
    rapport.append(Paragraph(100*'=', styles['Center']))
    rapport.append(Spacer(0*mm, 5*mm))  # (width, height)
    rapport.append(Paragraph('Localisation du centre : ' + centre, styles['Center']))
    rapport.append(Paragraph('Date du rapport : ' + date_now.strftime("%Y/%m/%d"), styles['Center']))
    rapport.append(Paragraph('Nom du traiteur d"image : ' + NomTraiteur, styles['Center']))
    rapport.append(Spacer(0*mm, 18*mm))  # (width, height)

    ########################################### Résumé ##################################################"

    rapport.append(Paragraph('<b><u>Résumé :', styles['Left']))
    rapport.append(Paragraph(140*'-', styles['Center']))

    # Informations générales

    rapport.append(Paragraph('<u>Informations générales :', styles['Indent']))
    rapport.append(Spacer(0*mm, 2*mm))  # (width, height)
    for cle, info in dGen.iteritems():
        rapport.append(Table([[Paragraph(cle + ' : ', styles['Indent2']), Paragraph(info, styles['Left'])]]))

    rapport.append(Table([[Paragraph('Machine IRM :', styles['Indent2']), Paragraph(MarqueMachine + ' ' + ChampMachine + 'T', styles['Left'])]]))

    if 'Patient' in dGen['Etiologie']:
        rapport.append(Table([[Paragraph('Date du traumatisme : ', styles['Indent2']), Paragraph('Inconnu', styles['Left'])]]))

    rapport.append(Spacer(0*mm, 2*mm))  # (width, height)
    rapport.append(Paragraph(140*'-', styles['Center']))

    # Résultat du controle qualité

    rapport.append(Paragraph('<u>Controle qualité :', styles['Indent']))
    for seq, chemin in dSeq.iteritems():
        if len(chemin) > 0 and tableau[dernierIdx]['CQ'][0]['Statut'][0][didSeq[seq]][0].encode('utf-8') == 'OK':
            resultat_protocol = Table([[Paragraph('Respect du protocole ' + seq + ' : ', styles['Indent2']), logo_OK]])
            resultat_protocol.hAlign = 'CENTER'
            rapport.append(resultat_protocol)
        else:
            if len(chemin) > 0:
                resultat_protocol = Table([[Paragraph('Respect du protocole ' + seq + ' : ', styles['Indent2']), logo_refus,tableau[dernierIdx]['CQ'][0]['Statut'][0][didSeq[seq]][0].encode('utf-8')]])
            else:
                resultat_protocol = Table([[Paragraph('Respect du protocole ' + seq + ' : ', styles['Indent2']), logo_refus]])
            resultat_protocol.hAlign = 'CENTER'
            rapport.append(resultat_protocol)

    if tableau[dernierIdx]['Statut'][0][0].encode('utf-8') == 'OK':
        resultat_qualite = Table([[Paragraph('<b>Avis final sur le volontaire : ', styles['Indent2']), logo_OK]])
    else:
        resultat_qualite = Table([[Paragraph('<b>Avis final sur le volontaire : ', styles['Indent2']), logo_refus]])

    resultat_qualite.hAlign = 'CENTER'

    rapport.append(resultat_qualite)
    ligneresult = rapport.index(resultat_qualite)
    rapport.append(Spacer(0*mm, 2*mm))  # (width, height)
    rapport.append(Paragraph(80*'=', styles['Center']))
    rapport.append(Spacer(0*mm, 2*mm))  # (width, height)
    rapport.append(PageBreak())
    ########################################### Analyse des champs DICOM #################################

    rapport.append(Paragraph('<b><u>Analyse des champs DICOM:', styles['Left']))
    rapport.append(Paragraph(140*'-', styles['Center']))
    rapport.append(Spacer(0*mm, 2*mm)) # (width, height)

    # Paramètres

    for seq, Chemin in dSeq.iteritems():
        rapport.append(Paragraph('<u>' + seq, styles['Indent']))
        rapport.append(Spacer(0*mm, 2*mm)) # (width, height)
        if len(Chemin) > 0 and len(glob.glob(Path + Chemin + "/Différences_Champs_Dicom*")) > 0:
            Dicomdiff = sio.loadmat(glob.glob(Path + Chemin + "/Différences_Champs_Dicom*")[0])
            #Dicomdifbis = Dicomdiff['Differencebis']
            rapport.append(Paragraph(str(Dicomdiff['NbC'][0][0]) + ' champs DICOM comparés', styles['Indent2']))
            NBmanquename = Dicomdiff['NBmanquename']
            if len(NBmanquename) > 0 and len(NBmanquename[0]) > 10:
                NBmanquename = NBmanquename[0]
                rapport.append(Paragraph(str(len(NBmanquename)) + ' champs DICOM mineurs, manquants par rapport à la référence du centre' , styles['Indent2']))
                ListeChampsManquant = []
                for champs in NBmanquename:
                    ListeChampsManquant.append(Paragraph('-' + champs[0].encode('utf-8'), styles['Center']))
                    if len(ListeChampsManquant) > 2:
                        NomChampsManquant = Table([ListeChampsManquant])
                        rapport.append(NomChampsManquant)
                        ListeChampsManquant = []

            rapport.append(Spacer(0*mm, 2*mm))  # (width, height)
    #        rapport.append(Paragraph(140*'-',styles['Center']))
    #        rapport.append(Spacer(0*mm, 2*mm)) # (width, height)

            # #Alertes Liste Noire

            if len(tableau[dernierIdx]['CQ'][0]['Anomalie'][0][didSeq[seq]]) > 0 and tableau[dernierIdx]['CQ'][0]['Statut'][0][didSeq[seq]][0].encode('utf-8') != 'OK':
                rapport.append(Paragraph('<u>Alertes :', styles['Indent2Red']))
                rapport.append(Spacer(0*mm, 2*mm)) # (width, height)
                TabEntete = Table([[Paragraph('<b>Champs DICOM', styles['Center']), Paragraph('<b>Référence', styles['Center']), Paragraph('<b>Sujet', styles['Center'])]])
                rapport.append(TabEntete)

                for i in range(len(tableau[dernierIdx]['CQ'][0]['Anomalie'][0][didSeq[seq]])):
                    try:
                        tab2 = tableau[dernierIdx]['CQ'][0]['Anomalie'][0][didSeq[seq]][i][1][0].encode('utf-8')
                    except AttributeError:
                        tab2 = str(tableau[dernierIdx]['CQ'][0]['Anomalie'][0][didSeq[seq]][i][1][0][0])
                    try:
                        tab3 = tableau[dernierIdx]['CQ'][0]['Anomalie'][0][didSeq[seq]][i][2][0].encode('utf-8')
                    except AttributeError:
                        tab3 = str(tableau[dernierIdx]['CQ'][0]['Anomalie'][0][didSeq[seq]][i][2][0][0])

                    ListeNoire = Table([[Paragraph(tableau[dernierIdx]['CQ'][0]['Anomalie'][0][didSeq[seq]][i][0][0].encode('utf-8'), styles['Center']), Paragraph(tab2, styles['Center']), Paragraph(tab3, styles['CenterRed'])]])
                    rapport.append(ListeNoire)
                    ListeNoire = []
                    del tab2, tab3
            if len(tableau[dernierIdx]['CQ'][0]['DirEncPhase'][0][didSeq[seq]][0]) > 0 and tableau[dernierIdx]['CQ'][0]['DirEncPhase'][0][didSeq[seq]][0].encode('utf-8')=='Erreur' and tableau[dernierIdx]['CQ'][0]['Statut'][0][didSeq[seq]][0].encode('utf-8')!='OK':
                    rapport.append(Paragraph('La direction d"encodage de phase de cette séquence de diffusion ne correspond pas au protocole demandé', styles['Indent2Red']))

            if tableau[dernierIdx]['CQ'][0]['Statut'][0][didSeq[seq]][0].encode('utf-8') == 'OK':
                    rapport.append(Paragraph('Pas de différences éliminatoires par rapport au protocole demandé', styles['Indent2']))

        else:
            rapport.append(Paragraph('Pas d"analyse de champs DICOM disponible', styles['Indent2']))
            if len(Chemin) == 0:
                rapport.append(Paragraph('Sujet référence ou données manquantes', styles['Indent2']))


        rapport.append(Spacer(0*mm, 2*mm))  # (width, height)
        rapport.append(Paragraph(140*'-', styles['Center']))
    rapport.append(PageBreak())
    # Critères qualité
    if os.path.exists(Path + centre + '/Sujets/' + sujet + '/QC_crit.png'):
        rapport.append(PageBreak())
        rapport.append(Paragraph(80*'=', styles['Center']))
        rapport.append(Spacer(0*mm, 2*mm))  # (width, height)
        QCcrit = Image(Path + centre + '/Sujets/' + sujet + '/QC_crit.png', 200*mm, 175*mm)
        QCcrit.hAlign = 'CENTER'
        rapport.append(Paragraph('<b>Critères qualité inter-sujets & intra-centre', styles['Left']))
        rapport.append(Paragraph(140*'-', styles['Center']))
        rapport.append(Spacer(0*mm, 1*mm)) # (width, height)
        rapport.append(QCcrit)

        rapport.append(Spacer(0*mm, 2*mm)) # (width, height)
        rapport.append(Paragraph('<b>SNR</b> : Valeur moyenne de l"image dans le cerveau divisée par l"écart type de l"image dans l"air (hors de la tete). <u>Les valeurs hautes sont préférables</u>.',styles['Left']))
        rapport.append(Paragraph('<b>CNR</b> : Différence d"intensité moyenne entre la matière grise et la matière blanche divisée par l"ecart type de l"image dans l"air. <u>Les valeurs hautes sont préférables</u>.',styles['Left']))
        rapport.append(Paragraph('<b>EFC</b> : L"EFC utilise l"entropie de Shannon de l"intensités des voxels comme une indication du ghosting et du flou induit par les mouvements de la tete. <u>Les valeurs faibles sont préférables</u>.[Atkinson 1997]',styles['Left']))
        rapport.append(Paragraph('<b>FBER</b> : Définit comme l"énergie moyenne des valeurs à l"intérieur de la tete par rapport à l"extérieur de la tete. <u>Les valeurs hautes sont préférables</u>.',styles['Left']))



    #######################################################Analyse DTI########################################################################

    if len(tableau[dernierIdx]['EddyCorr'][0]['Statut'][0][0])>0 and 'OK' in tableau[dernierIdx]['EddyCorr'][0]['Statut'][0][0][0].encode('utf-8'):
        rapport.append(PageBreak())
        rapport.append(Paragraph(80 * '=', styles['Center']))
        rapport.append(Paragraph('<b><u>Qualité des données:', styles['Left']))
        rapport.append(Spacer(0 * mm, 2 * mm))  # (width, height)
        rapport.append(Paragraph(140 * '-', styles['Center']))
        # Mesures globales
        if os.path.exists(Path + centre + '/Sujets/' + sujet + '/DTI_Pretrt/MesureQualite.mat'):

            test2 = sio.loadmat(Path + centre + '/Sujets/' + sujet + '/DTI_Pretrt/MesureQualite.mat')
            Mesures = test2['Mesures']
            if Mesures[5][0] < pow(Mesures[8][0], 1.0/3.0):
                meanmvmtimresult = logo_OK
            else:
                meanmvmtimresult = logo_refus
                #rapport.remove(resultat_qualite)
                #resultat_qualite = Table([[Paragraph('<b>Qualités des données : ', styles['Indent2']), logo_refus]])
                #resultat_qualite.hAlign = 'CENTER'
                #rapport.insert(ligneresult, resultat_qualite)
            if Mesures[6][0] < 1.5*pow(Mesures[8][0], 1.0/3.0):
                maxmvmtimresult = logo_OK
            else:
                maxmvmtimresult = logo_refus
                #rapport.remove(resultat_qualite)
                #resultat_qualite = Table([[Paragraph('<b>Qualités des données : ', styles['Indent2']), logo_refus]])
                #resultat_qualite.hAlign = 'CENTER'
                #rapport.insert(ligneresult, resultat_qualite)
            meanmvmtimresult.hAlign = 'CENTER'
            maxmvmtimresult.hAlign = 'CENTER'
            TabMesures = [[Paragraph('<b>Paramètre : ', styles['Left']), Paragraph('<b>Valeur', styles['Left']), Paragraph('<b>Normes', styles['Center']), Paragraph('<b>Résultat', styles['Left'])],
                        [Paragraph('Volume cérébral : ', styles['Left']), Paragraph('{:.2e} mm<sup>3'.format(int(Mesures[4][0])), styles['Left']), Paragraph('', styles['Center']), Paragraph('', styles['Center'])],
                        [Paragraph('MD Matière grise : ', styles['Left']), Paragraph('{:.2e} mm<sup>2</sup>/s'.format(Mesures[0][0]), styles['Left']), Paragraph('', styles['Center']), Paragraph('', styles['Center'])],
                        [Paragraph('MD Matière blanche : ', styles['Left']), Paragraph('{:.2e} mm<sup>2</sup>/s'.format(Mesures[1][0]), styles['Left']), Paragraph('', styles['Center']), Paragraph('', styles['Center'])],
                        [Paragraph('FA Matière grise : ', styles['Left']), Paragraph('{:.2e}'.format(Mesures[2][0]), styles['Left']), Paragraph('', styles['Center']), Paragraph('', styles['Center'])],
                        [Paragraph('FA Matière blanche : ', styles['Left']), Paragraph('{:.2e}'.format(Mesures[3][0]), styles['Left']), Paragraph('', styles['Center']), Paragraph('', styles['Center'])],
                        [Paragraph('Mouvement moyen dans la diffusion :', styles['Left']), Paragraph('{:.3f}'.format(Mesures[5][0]), styles['Left']), Paragraph('< {:.1f}'.format(pow(Mesures[8][0], 1.0/3.0)), styles['Center']), meanmvmtimresult],
                        [Paragraph('Mouvement maximal dans la diffusion :', styles['Left']), Paragraph('{:.3f}'.format(Mesures[6][0]), styles['Left']), Paragraph('< {:.1f}'.format(1.5*pow(Mesures[8][0], 1.0/3.0)), styles['Center']), maxmvmtimresult]]
            Tabmesures = Table(TabMesures)
            rapport.append(Tabmesures)

        else:
            rapport.append(Paragraph('<b>Informations sur les mesures globales manquantes', styles['Indent2']))

        if len(tableau[dernierIdx]['CQ'][0][0][didSeq['DTI']]['NbCorrCoupe']) > 0:
            NbCCoupe = tableau[dernierIdx]['CQ'][0][0][didSeq['DTI']]['NbCorrCoupe'][0][0]
        else:
            NbCCoupe = 0
        NbCoupe = tableau[dernierIdx]['CQ'][0][0][didSeq['DTI']]['NbCoupe'][0][0]
        NbVol = tableau[dernierIdx]['CQ'][0][0][didSeq['DTI']]['NbVol'][0][0]
        NbVolSupp = tableau[dernierIdx]['CQ'][0][0][didSeq['DTI']]['VolSupp']

        n = 0
        if os.path.exists(Path + centre + '/Sujets/' + sujet + '/DTI_Pretrt/DTI_esc.nii.eddy_outlier_report'):
            n = sum(1 for _ in open(Path + centre + '/Sujets/' + sujet + '/DTI_Pretrt/DTI_esc.nii.eddy_outlier_report'))
            NbCorrCoupe =  n # pour la version suivante ou on force la correction eddy
            #NbCorrCoupe = NbCCoupe + n
        else:
            NbCorrCoupe = NbCCoupe

        TxCoupe = NbCorrCoupe/float(int(NbCoupe)*int(NbVol))
        TxVol = len(NbVolSupp)/float(int(NbVol))
        if TxCoupe < SeuilCoupe:
            txcoupeimresult = logo_OK
        else:
            txcoupeimresult = logo_refus
            #rapport.remove(resultat_qualite)
            #resultat_qualite = Table([[Paragraph('<b>Qualités des données : ', styles['Indent2']), logo_refus]])
            #resultat_qualite.hAlign = 'CENTER'
            #rapport.insert(ligneresult, resultat_qualite)
        if TxVol < SeuilVol:
            txvolimresult = logo_OK
        else:
            txvolimresult = logo_refus
            #rapport.remove(resultat_qualite)
            #resultat_qualite = Table([[Paragraph('<b>Qualités des données : ', styles['Indent2']), logo_refus]])
            #resultat_qualite.hAlign = 'CENTER'
            #rapport.insert(ligneresult, resultat_qualite)
        txcoupeimresult.hAlign = 'CENTER'
        txvolimresult.hAlign = 'CENTER'
        rapport.append(Spacer(0*mm, 1*mm))  # (width, height)
        TabCorr = [[Paragraph('Taux de coupes corrigées  : ', styles['Left']), Paragraph(str(TxCoupe)+' (' + str(NbCorrCoupe) + '/' + str(int(NbCoupe)*int(NbVol)) + ')', styles['Left']), Paragraph(str(SeuilCoupe), styles['Center']), txcoupeimresult],
                   [Paragraph('Taux de directions supprimées : ', styles['Left']), Paragraph(str(TxVol)+' (' + str(len(NbVolSupp)) + '/' + str(int(NbVol))+')', styles['Left']), Paragraph(str(SeuilVol), styles['Center']), txvolimresult]]

        Tabcorr = Table(TabCorr)
        rapport.append(Tabcorr)

        rapport.append(PageBreak())

        #Figure récapitulative QC
        rapport.append(Paragraph(140*'-', styles['Center']))
        rapport.append(Spacer(0*mm, 1*mm))  # (width, height)
        if os.path.exists(Path + centre + '/Sujets/' + sujet + '/DTI_Pretrt/QC.png'):
            QC = Image(Path + centre + '/Sujets/' + sujet + '/DTI_Pretrt/QC.png', 195*mm, 190*mm)
            QC.hAlign = 'CENTER'
            rapport.append(QC)

        rapport.append(Spacer(0*mm, 2*mm))  # (width, height)
        rapport.append(Paragraph('<b>Signal moyen des coupes axiales</b> : Intensité moyenne des voxels intra-cérébraux de chaque coupe pour chaque direction. Les barres d"erreur bleues et vertes correspondent respectivement à 3 et 5  fois l"écart types calculé sur toutes les directions.',styles['Left']))
        rapport.append(Paragraph('<b>Direction gradients de diffusion</b> : Directions des gradients de diffusion sur une sphère. Les directions supprimées sont affichées en rouge.',styles['Left']))
        rapport.append(Paragraph('<b>Résidu du modèle de tenseur</b> Différence cumulée des intensités de voxels intra-cérébraux entre l"image de diffusion et l"estimation réalisée à partir du tenseur de diffusion pour chaque coupe et chaque direction.',styles['Left']))
        rapport.append(Paragraph('<b>Le mouvement total</b> : Mesure du mouvement dans chaque volume calculée à partir du déplacement de chaque voxel, en moyennant les carrés de ces déplacements à travers les voxels intracérébraux et en calculant la racine carrée de ce résultat. Le mouvement absolu représente le mouvement(Root Mean Squared) par rapport à la première direction. Le mouvement relatif représente le mouvement par rapport à la direction précédente',styles['Left']))
        rapport.append(Paragraph('<b>Rotation et Translation</b> : Paramètres de translation et de rotations utilisés dans la coregistration des différentes directions sur la première direction.',styles['Left']))


        if os.path.exists(Path + centre + '/Sujets/' + sujet + '/DTI_Pretrt/QC.png'):
            rapport.append(PageBreak())
            rapport.append(Paragraph(140 * '-', styles['Center']))
            rapport.append(Spacer(0 * mm, 1 * mm))  # (width, height)
            QC2 = Image(Path + centre + '/Sujets/' + sujet + '/DTI_Pretrt/QC2.png', 200*mm, 150*mm)
            QC2.hAlign = 'CENTER'
            rapport.append(QC2)
            rapport.append(Spacer(0 * mm, 2 * mm))  # (width, height)
            rapport.append(Paragraph('<b>Intensité moyenne des coupes paires et impaires</b> : Distribution des intensités moyennes pour les coupes paires et les coupes impaires.',styles['Left']))
            rapport.append(Paragraph('<b>Histogramme sphérique de la direction de la V1 pour tous les voxels du cerveau</b> : Le nombre de voxel partageant la même direction principale V1 est représenté sur une sphére',styles['Left']))

        # Figure récapitulative par centre FAMDROI
        if os.path.exists(Path + centre + '/Sujets/' + sujet + '/FAMDROI.png'):
            rapport.append(PageBreak())
            rapport.append(Paragraph(80*'=', styles['Center']))
            rapport.append(Spacer(0*mm, 2*mm))  # (width, height)
            FAMDROI = Image(Path + centre + '/Sujets/' + sujet + '/FAMDROI.png', 215*mm, 215*mm)
            FAMDROI.hAlign = 'CENTER'
            rapport.append(Paragraph('<b>Mesures inter-sujets & intra-centre', styles['Left']))
            rapport.append(Paragraph(140*'-', styles['Center']))
            rapport.append(Spacer(0*mm, 1*mm))  # (width, height)
            rapport.append(FAMDROI)

        #         rapport.append(Spacer(0*mm, 2*mm))  # (width, height)

        ##        rapport.append(PageBreak())
        ##        rapport.append(Paragraph('<b>Index des ROI',styles['Center']))
        ##        labelAtlas= sio.loadmat(Path + "ROIsAtlas2.mat")
        ##        Tablabel=[[Paragraph('<b>Matière Grise',styles['Right']),Paragraph('',styles['Center']),Paragraph('<b>Matière Blanche',styles['Right'])],
        ##                  [Paragraph('<b>N° de la ROI',styles['Left']),Paragraph('<b>Nom de la ROI',styles['Left']),Paragraph('<b>N° de la ROI',styles['Left']),Paragraph('<b>Nom de la ROI',styles['Left'])]]
        ##        rapport.append(Table(Tablabel))
        ##        for i in range(len(labelAtlas['ROIMG'])):
        ##            try:
        ##                labelROIMB=labelAtlas['ROIMB'][i][0][0].encode('utf-8')
        ##                numlabelMB= i+ 1 + len(labelAtlas['ROIMG'])
        ##            except:
        ##                labelROIMB=''
        ##                numlabelMB=''
        ##            Tablabel=[[Paragraph(str(i+1),styles['Left']),Paragraph(labelAtlas['ROIMG'][i][0][0].encode('utf-8'),styles['Left']),Paragraph(str(numlabelMB),styles['Left']),Paragraph(labelROIMB,styles['Left'])]]
        ##            rapport.append(Table(Tablabel))

    page.build(rapport, canvasmaker=PageNumCanvas)


if __name__ == '__main__':
    dernierIdx = float(sys.argv[1])
    generation_rapport(dernierIdx)
