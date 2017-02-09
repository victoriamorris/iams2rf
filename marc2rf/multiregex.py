#!/usr/bin/env python
# -*- coding: utf8 -*-

"""Classes for making multiple regex replacements in the Researcher Format transformation."""

import regex as re
import sys

__author__ = 'Victoria Morris'
__license__ = 'MIT License'
__version__ = '1.0.0'
__status__ = '4 - Beta Development'


class MultiRegex(object):
    simple = False
    regexes = ()

    def __init__(self):
        try: self._rx = re.compile('|'.join(self.regexes), flags=re.IGNORECASE)
        except:
            for r in self.regexes:
                try: re.compile(r)
                except:
                    print('Error in regex: {}'.format(str(r)))

    def sub(self, s):
        if not s or s is None: return ''
        return self._rx.sub(self._sub, s)

    def _sub(self, mo):
        try:
            for k, v in mo.groupdict().items():
                if v:
                    if k == 'AllElse':
                        return ''
                    if 'UUU' in str(k):
                        return bytes(str(k).replace('UUU', '\\' + 'u'), 'ascii').decode('unicode-escape')
                    try:                         
                        sub = getattr(self, k)
                        if callable(sub): return sub(mo)
                        else: return sub
                    except: return str(k)              
                    
        except:
            print('\nError MR: {0}\n'.format(str(sys.exc_info())))


# List of MultiRegex classes:
#   Abbreviations
#   Genres
#   PlaceNamesUK
#   PlaceNamesUS
#   PlaceNamesAustralia
#   PlaceNamesBrazil
#   PlaceNamesCanada
#   PlaceNamesNewZealand
#   PlaceNamesOther
#   PlaceNamesAccents
#   Relators


class Abbreviations(MultiRegex):
    simple = True
    regexes = (
        r'(?P<January>^jan(uary)?\.*$)',
        r'(?P<February>^feb(ruary)?\.*$)',
        r'(?P<March>^m(ar|rz)(ch)?\.*$)',
        r'(?P<April>^apr(il)?\.*$)',
        r'(?P<June>^june?\.*$)',
        r'(?P<July>^july?\.*$)',
        r'(?P<August>^aug(ust)?\.*$)',
        r'(?P<September>^sept?(ember)?\.*$)',
        r'(?P<October>^o[ck]t(ober)?\.*$)',
        r'(?P<November>^nov(ember)?\.*$)',
        r'(?P<December>^de[cz](ember)?\.*$)',
        r'(?P<Monday>^mon(day)?s?\.*$)',
        r'(?P<Tuesday>^tues?(day)?s?\.*$)',
        r'(?P<Wednesday>^wed(ne)?s?(day)?s?\.*$)',
        r'(?P<Thursday>^thur?s?(day)?s?\.*$)',
        r'(?P<Friday>^fri(day)?s?\.*$)',
        r'(?P<Saturday>^sat(urday)?s?\.*$)',
        r'(?P<Sunday>^sun(day)?s?\.*$)',
        r'(?P<Abbildung>^abb(ildung)?\.*$)',    # German, illustration, figure
        r'(?P<Abdruck>^abdr(uck)?\.*$)',        # German, impression, print, reproduction
        r'(?P<Abhandlung>^abh(andlung)?\.*$)',  # German, treatise
        r'(?P<AbkUUU00FCrzung>^abk(.rzung)?\.*$)',  # German, abbreviation
        r'(?P<Abschrift>^abschr(ift)?\.*$)',    # German, reprint, copy
        r'(?P<Abteilung>^abt(eilung)?\.*$)',    # German
        r'(?P<approximately>^(ca|approx)\.*$)',
        r'(?P<Auflage>^aufl(age)?\.*$)',     # German, edition
        r'(?P<Ausgabe>^ausg(abe)?\.*$)',     # German, edition
        r'(?P<augmented>^aug(mented)\.*$)',
        r'(?P<BUUU00E4ndchen>^b(aen)?dche?n\.*$)',   # German
        r'(?P<BUUU00E4nde>^b(ae?n)?de\.*$)',    # German
        r'(?P<Band>^b(an)?d\.*$)',              # German, volume
        r'(?P<Bearbeitung>^bearb(eitung)?\.*$)',    # German, arrangement
        r'(?P<Beiheft>^beih(eft)?\.*$)',        # German, supplement
        r'(?P<Beispiel>^beisp(iel)?\.*$)',      # German, example
        r'(?P<beziehungsweise>^be?z(iehungs)?w(eise)?\.*$)',    # German, respectively; or, or else; more specifically
        r'(?P<bibliography>^bibl(iog)?(raphy)?\.*$)',
        r'(?P<books>^bo*ks\.*$)',
        r'(?P<book>^bo*k\.*$)',
        r'(?P<Buchhandler>^buchh(andler)?\.*$)',    # German, bookseller
        r'(?P<CDs>^cd-?(rom)?s\.*$)',
        r'(?P<CD>^cd-?(rom)?\.*$)',
        r'(?P<chiefly>^chiefle*y\.*$)',        
        r'(?P<cm>^cm\.*$)',
        r'(?P<coloured>^colo+u?red\.*$)',
        r'(?P<colour>^col(o+u?r|eur)?\.*$)',
        r'(?P<columns>^col(umn)?s\.*$)',
        r'(?P<corrected>^corr(ected)?\.*$)',
        r'(?P<cover>^couv(erture)?\.*$)',
        r'(?P<deel>^de*l\.*$)',     # Dutch
        r'(?P<diagrams>^diagra?m?s*\.*$)',
        r'(?P<dopolnennoe>^dop(ol)?(nennoe)?\.*$)',  # Russian
        r'(?P<DVDs>^dvd-?(rom)?s\.*$)',
        r'(?P<DVD>^dvd-?(rom)?\.*$)',
        r'(?P<UUU00E9dition>^[\u00e9\u00C9]d(ition)?\.*$)',     # édition
        r'(?P<edition>^ed(ition)?\.*$)',
        r'(?P<Einleitung>^einl(eitung)?\.*$)',  # German, introduction
        r'(?P<ekdosi>^ekd(osi)?\.*$)',          # Greek
        r'(?P<engraved>^engr(aved)?\.*$)',
        r'(?P<enlarged>^enl(arged)?\.*$)',
        r'(?P<erweiterte>^erw(eit)?(erte)?\.*$)',   # German
        r'(?P<fascicule>^fasc(icule)?\.*$)',    # French
        r'(?P<facsimiles>^fa(cs|sc)(im)?(ile)?s\.*$)',
        r'(?P<facsimile>^fa(cs|sc)(im)?(ile)?\.*$)',
        r'(?P<feet>^f[e]*t\.*$)',
        r'(?P<figures>^fig(ures)?s*\.*$)',
        r'(?P<folded>^(ofld|fold(ed)?)\.*$)',
        r'(?P<folio>^fol[io.]*\.*$)',
        r'(?P<folios>^fol[io.]*s\.*$)',
        r'(?P<frames>^fr(ame)?s*\.*$)',
        r'(?P<frontispiece>^front(\.|is)(piece)?\.*$)',
        r'(?P<gedruckt>^gedr(uckt)?\.*$)',      # German, printed
        r'(?P<Gegenwart>^gegenw(art)?\.*$)',    # German, present time
        r'(?P<genealogical>^geneal(ogical)?\.*$)',
        r'(?P<geological>^geol(og)?(ical)?\.*$)',
        r'(?P<garren>^g(arre)?n\.*$)',          # Basque, nth
        r'(?P<Handbuch>^h(an)?db(uch)?\.*$)',   # German, handbook, manual
        r'(?P<hardback>^h(ard)?b(ac)?k\.*$)',
        r'(?P<Hefte>^he*fte\.*$)',              # German
        r'(?P<Heft>^he*ft\.*$)',                # German
        r'(?P<Herausgeber>^he?r(au)?sg(eber)?\.*$)',    # German, editor
        r'(?P<illustrations>^a?il+u?s?(tration.*)?s?\.*$)',
        r'(?P<impression>^impr(ession)?\.*$)',
        r'(?P<including>^incl?(uding)?\.*$)',
        r'(?P<introduction>^introd(uction)?\.*$)',
        r'(?P<ispravlennoe>^ispr(avl)?(ennoe)?\.*$)',   # Russian
        r'(?P<izdaniye>^izd(aniye)?\.*$)',      # Russian
        r'(?P<Jahreszahl>^j(ahres)?z(ah)?l\.*$)',       # German, date, year
        r'(?P<jaargang>^jaarg(ang)?\.*$)',      # Dutch
        r'(?P<Jahrgang>^jahrg(ang)?\.*$)',      # German
        r'(?P<Jahrhundert>^j(ahr)?h(undert)?\.*$)',     # German, century
        r'(?P<knjiga>^knj(iga)?\.*$)',          # Croatian
        r'(?P<mahadurah>^mahad(urah)?\.*$)',    # Hebrew
        r'(?P<manuscript>^m(ss*|anuscripts?)\.*$)',
        r'(?P<microfiche>^micr[io]-*fiches*\.*$)',
        r'(?P<microfilm>^micr[io]-*film*\.*$)',
        r'(?P<minutes>^min(ute)?s\.*$)',
        r'(?P<Mitarbeiter>^mitarb(eiter)?\.*$)',    # German, collaborator
        r'(?P<Mitwirkung>^mitw(irkung)?\.*$)',      # German, cooperation
        r'(?P<mm>^mm\.*$)',
        r'(?P<music>^mus(ic)?\.*$)',
        r'(?P<Nachricht>^nachr(icht)?\.*$)',    # German, communication, report, notice
        r'(?P<Nachwort>^nachw(ort)?\.*$)',      # German, concluding remarks, epilogue
        r'(?P<nakladateUUU0142stvUUU00ed>^nakl(ad)?(ate)?\.*$)',      # Czech, nakladatełství
        r'(?P<Neudruck>^neudr(uck)?\.*$)',      # German, reprint
        r'(?P<nouvelle>^nouv(elle)?\.*$)',      # French
        r'(?P<numbers>^n-*(o|ro?|um+b?ero?)s*\.*$)',
        r'(?P<oblong>^obl(ong)?\.*$)',
        r'(?P<Originalausgabe>^Originalausg(abe)?\.*$)',        # German
        r'(?P<pages>^pp+(age)?s*\.*$)',
        r'(?P<paperback>^p(aper)?b(ac)?k\.*$)',
        r'(?P<parts>^p(ar)?t\.*$)',
        r'(?P<patippu>^pat(ippu)?\.*$)',        # Russian
        r'(?P<plates>^pl(at)?e?s*\.*$)',
        r'(?P<poprawione>^popr(awione)?\.*$)',  # Polish, corrected       
        r'(?P<portraits>^portr?(ait)?s*\.*$)',
        r'(?P<reprinted>^re-*pr(int)?(ed)?\.*$)',
        r'(?P<revised>^rev(ised)?\.*$)',
        r'(?P<Sammelwerk>^s(ammel)?w(er)?k\.*$)',       # German, collected works
        r'(?P<Sammlung>^samml(ung)?\.*$)',              # German, collection, compilation, set
        r'(?P<Schriftleiter>^schriftl(eiter)?\.*$)',    # German, editor
        r'(?P<selfUUU002Dportraits>^self-?portr?(ait)?s*\.*$)',
        r'(?P<series>^ser(ies)?\.*$)',
        r'(?P<sheet>^sh\.*$)',
        r'(?P<stereograph>^stereo-?graph\.*$)',
        r'(?P<sound>^s(oun)?d\.*$)',
        r'(?P<Stimmbuch>^st(imm)?b(uch)?\.*$)',     # German, part book
        r'(?P<supplement>^suppl?(ement)?\.*$)',
        r'(?P<svazek>^sv(azek)?\.*$)',      # Czech
        r'(?P<tomes>^tome?s*\.*$)',
        r'(?P<undUUU0020soUUU0020weiter>^u(nd)?\s*so?\s*w(eiter)?\.*$)',    # German, and so forth, etc.
        r'(?P<unnumbered>^un-?numbered\.*$)',
        r'(?P<updated>^upd(ated)?\.*$)',
        r'(?P<uzupeUUU0142nione>^uzup(elnione)?\.*$)',  # Polish, uzupełnione
        r'(?P<Verfasser>^verf(asser)?\.*$)',        # German, composer, writer
        r'(?P<vergleich>^vergl(eich)?\.*$)',        # German, compare
        r'(?P<Verzeichnis>^verz(eichnis)?\.*$)',    # German, catalogue
        r'(?P<videodisc>^video-*disc\.*$)',
        r'(?P<volumes>^vol?(ume)?s*\.*$)',
        r'(?P<Vorwort>^vorw(ort)?\.*$)',    # German, foreword
        r'(?P<vydUUU00E1nUUU00ED>^vyd(ani)?\.*$)',      # Czech, vydání
        r'(?P<vypusk>^vyp(usk)?\.*$)',      # Russian
        r'(?P<wydanie>^wyd(anie)?\.*$)',    # Polish
        r'(?P<years>^y(ea)?rs\.*$)',
        r'(?P<year>^y(ea)?r\.*$)',
        r'(?P<Zeitschrift>^z(ei)?tschr(ift)?\.*$)',     # German, periodical
        r'(?P<Zeitung>^z(ei)?t(un)?g\.*$)',  # German, newspaper
        r'(?P<zeszyt>^zesz(yt)?\.*$)',      # Polish
        r'(?P<zvezek>^zv(ezek)?\.*$)',      # Slovenian, volumes
        )            


class Genres(MultiRegex):
    simple = True
    regexes = (
        r'(?P<TranslationsInto>.*\btranslations.*?\binto.*?\b([a-z]+).*)',
        r'(?P<TranslationsFrom>.*\btranslations.*?\bfrom.*?\b([a-z]+).*)',
        r'(?P<EarlyWorks>.*\bearly.*?\bworks.*?\bto.*?\b([0-9]+).*)',
        r'(?P<YoungUUU0020adultUUU0020fiction>.*\byoung.*?\badult.*?\b(stor|fiction).*)',
        r'(?P<YearbookUUU0020>.*\byear[\-\s]*book.*?\b.*)',
        r'(?P<Woodcut>.*\bwood[\-\s]*cut.*)',
        r'(?P<WesternUUU0020fiction>.*\bwestern.*?\b(stor|fiction).*)',
        r'(?P<Western>.*\bwestern.*)',
        r'(?P<WebUUU0020site>.*\bweb[\-\s]*site.*)',
        r'(?P<Watercolour>.*\bwater[\-\s]*colou*r.*)',
        r'(?P<VocalUUU0020score>.*\bvocal.*?\bscore.*)',
        r'(?P<UtopianUUU0020fiction>.*\butopia.*?\b(stor|fiction).*)',
        r'(?P<Utopia>.*\butopia.*)',
        r'(?P<UrbanUUU0020fiction>.*\burban.*?\b(stor|fiction).*)',
        r'(?P<Urban>.*\burban.*)',        
        r'(?P<Videorecording>.*\bvideo.*?\brecording.*)',
        r'(?P<TrueUUU0020crime>.*\btrue.*?\bcrime.*)',
        r'(?P<TrueUUU0020adventure>.*\btrue.*?\badventure.*)',
        r'(?P<Treaty>.*\btreaty.*)',
        r'(?P<Treatise>.*\btreatise.*)',
        r'(?P<TravelUUU0020guide>.*\btravel[\-\s]*guide.*)',
        r'(?P<Travelogue>.*\btravelog.*)',
        r'(?P<Travel>.*\btravel\b.*)',
        r'(?P<Transparency>.*\btransparenc.*)',
        r'(?P<Transcript>.*\btranscript.*)',
        r'(?P<TrialUUU0020orUUU0020litigation>.*\b(trial|litigation)s*\b.*)',
        r'(?P<Trailer>.*\btrailer.*)',
        r'(?P<Tragicomedy>.*\btragi[\-\s]*comedy.*)',
        r'(?P<Tragedy>.*\btragedy.*)',
        r'(?P<ToyUUU0020orUUU0020moveableUUU0020book>.*\b(toy|move*able).*?\bbook.*)',
        r'(?P<Toy>.*\btoy.*)',
        r'(?P<Thriller>.*\bthriller.*)',
        r'(?P<Thesaurus>.*\bthesaur.*)',
        r'(?P<Textbook>.*\btext[\-\s]*book.*)',
        r'(?P<TechnicalUUU0020report>.*\btechnical.*?\breport.*)',
        r'(?P<TechnicalUUU0020drawing>.*\btechnical.*?\bdrawing.*)',
        r'(?P<TeachingUUU0020material>.*\bteach.*?\b(material|resource|guide).*)',
        r'(?P<Taxonomy>.*\btaxonom.*)',
        r'(?P<TallUUU0020tale>.*\btall.*?\btale.*)',
        r'(?P<SuspenseUUU0020fiction>.*\bsuspense.*?\b(stor|fiction).*)',
        r'(?P<Suspense>.*\bsuspense.*)',
        r'(?P<SurveyUUU0020ofUUU0020literature>.*\bsurvey.*?\bliterature.*)',
        r'(?P<StyleUUU0020guide>.*\bstyle.*?\b(guide|manual).*)',
        r'(?P<StudyUUU0020guide>.*\bstudy.*?\b(gu*ide|aid).*)',
        r'(?P<StreamingUUU0020video>.*\bstream.*?\bvideo.*)',
        r'(?P<StoryUUU0020inUUU0020rhyme>.*\b(stor(y|ies)|fiction).*?\brhyme.*)',
        r'(?P<StatuteUUU0020orUUU0020ordinance>.*\b(statute|ord(in|onn)ance).*)',
        r'(?P<Statistics>.*\bstatistics.*)',
        r'(?P<StandardUUU0020orUUU0020specification>.*\b(standard|specification).*)',
        r'(?P<SpyUUU0020story>.*\b(spy|espionage).*?\b(stor|fiction).*)',
        r'(?P<SportUUU0020fiction>.*\bsport(s|ing).*?\b(stor|fiction).*)',
        r'(?P<Sport>.*\bsport(s|ing)?\b.*)',
        r'(?P<Speech>.*\bspeech.*)',
        r'(?P<Specimen>.*\bspecimen.*)',
        r'(?P<Source>.*\bsources*\b.*)',
        r'(?P<SoundUUU0020recording>.*\bsound.*?\brecord.*)',
        r'(?P<SoundUUU0020effect>.*\bsound.*?\beffect.*)',
        r'(?P<Sound>.*\bsound.*)',
        r'(?P<Song>.*\bsongs*\b.*)',
        r'(?P<Slide>.*\bslide.*)',
        r'(?P<ShortUUU0020story>.*\bshort.*?\bstor.*)',
        r'(?P<Sermon>.*\bsermon.*)',
        r'(?P<Series>.*\bseries.*)',
        r'(?P<SelfUUU002Dhelp>.*\bself.*?\bhelp.*)',
        r'(?P<Script>.*\bscript.*)',
        r'(?P<ScienceUUU0020fiction>.*\bsci(ence)?.*?\bfi(ction)?.*)',
        r'(?P<SatelliteUUU0020image>.*\bsatellite.*?\bimage.*)',
        r'(?P<ParanormalUUU0020romanceUUU0020fiction>.*\bparanormal.*?\b(roman(ce|s)|love).*?\b(stor|fiction).*)',
        r'(?P<RomanceUUU0020fiction>.*\b(roman(ce|s)|love).*?\b(stor|fiction).*)',
        r'(?P<Romance>.*\b(roman(ce|s)|love).*)',
        r'(?P<Review>.*\breview.*)',
        r'(?P<Reporting>.*\breport(ing|age).*)',
        r'(?P<AnnualUUU0020report>.*\bannual.*?\breport.*)',
        r'(?P<Report>.*\breport\b.*)',
        r'(?P<RemoteUUU0020sensingUUU0020image>.*\bremote.*?\bsensing.*?\bimage.*)',
        r'(?P<ReligiousUUU0020text>.*\breligious.*?\btext.*)',
        r'(?P<Rehearsal>.*\brehearsal.*)',
        r'(?P<Register>.*\bregister.*)',
        r'(?P<ReferenceUUU0020data>.*\breference.*?\bdata.*)',
        r'(?P<Reference>.*\breference.*)',
        r'(?P<Realism>.*\brealis(m|tic).*)',
        r'(?P<Realia>.*\brealia.*)',
        r'(?P<Reader>.*\breader.*)',
        r'(?P<QuotationUUU0020orUUU0020maxim>.*\b(quot(e|ation)|axim\b).*)',
        r'(?P<Quiz>.*\bquiz.*)',
        r'(?P<QuestionsUUU0020andUUU0020answers>.*\b(question.*?\banswer|q\b.*?\ba\b).*)',
        r'(?P<Puzzle>.*\bpuzzle.*)',
        r'(?P<PsychologicalUUU0020thriller>.*\bpsycho.*?\bthriller.*)',
        r'(?P<PseudoUUU0020documentary>.*\b((pseudo|fake|quasi|fiction|mock|style).*?\bdocumentar(y|ies)|documentar(y|ies).*?\b(pseudo|fake|quasi|fiction|mock|style)).*)',
        r'(?P<Psalter>.*\bpsalter.*)',
        r'(?P<ProverbUUU0020orUUU0020saying>.*\b(proverb|sayings*\b).*)',
        r'(?P<Prospectus>.*\bprospectus.*)',
        r'(?P<ProseUUU0020literature>.*\bprose.*)',
        r'(?P<Propaganda>.*\bpropaganda.*)',
        r'(?P<ProgrammedUUU0020text>.*\bprogrammed.*?\btext.*)',
        r'(?P<Proclamation>.*\bproclai*mation.*)',
        r'(?P<ProblemsUUU0020andUUU0020exercises>.*\b(problem|exercis).*)',
        r'(?P<Primer>.*\bprimers*\b.*)',
        r'(?P<PrayerUUU0020orUUU0020devotionalUUU0020book>.*\b(prayer|devotion).*)',
        r'(?P<Poster>.*\bposter.*)',
        r'(?P<Postcard>.*\bpost.*?\bcard.*)',
        r'(?P<PoetryorUUU0020verse>.*\b(poetry|verse).*)',
        r'(?P<Playbill>.*\bplay[\-\s]*bill.*)',
        r'(?P<PlanUUU0020orUUU0020view>.*\b(plan|view).*)',
        r'(?P<Picture>.*\bpicture.*)',
        r'(?P<PictorialUUU0020work>.*\bpicti*ori*al.*?\bwork.*)',
        r'(?P<PicaresqueUUU0020fiction>.*\bpicares.*?\b(stor|fiction).*)',
        r'(?P<Picaresque>.*\bpicares.*)',        
        r'(?P<PianoUUU0020score>.*\bpiano.*?\bscore.*)',
        r'(?P<Petition>.*\bpetition.*)',
        r'(?P<PersonalUUU0020narrative>.*\bperson.*?\bnarrat.*)',
        r'(?P<Periodical>.*\bperiodical.*)',
        r'(?P<Patent>.*\bpatent.*)',
        r'(?P<Parody>.*\b(parody|imitation|spoof).*)',
        r'(?P<ParanormalUUU0020fiction>.*\bparanormal.*?\b(stor|fiction).*)',
        r'(?P<Paranormal>.*\bparanormal.*)',
        r'(?P<Pamphlet>.*\bpamphlet.*)',
        r'(?P<OutlineUUU0020orUUU0020syllabus>.*\b(outline|syllab(i|us)).*)',
        r'(?P<Oration>.*\boration.*)',
        r'(?P<Opinion>.*\bopinion.*)',
        r'(?P<OnlineUUU0020systemUUU0020orUUU0020service>.*\bon-*line.*?\b(system|service).*)',
        r'(?P<Offprint>.*\boff-*print.*)',
        r'(?P<Obituary>.*\bobituar.*)',
        r'(?P<NumericUUU0020data>.*\bnumeric.*?\bdata.*)',
        r'(?P<Novel>.*\bnovel.*)',
        r'(?P<Nomenclature>.*\bnomenclature.*)',
        r'(?P<Newspaper>.*\bnewspaper.*)',
        r'(?P<Newsbook>.*\bnews[\-\s]*book.*)',
        r'(?P<Nature>.*\bnature.*)',
        r'(?P<NarrativeUUU0020nonUUU002Dfiction>.*\bnarrative.*?\bnon.*?\bfiction.*)',
        r'(?P<Mythopoeia>.*\bmythopoeia.*)',
        r'(?P<Mythology>.*\bmythology.*)',
        r'(?P<MusicalUUU0020arrangement>.*\b(music.*?\barrange|arrange.*?\bmusic).*)',
        r'(?P<SacredUUU0020music>.*\bsacred.*?\bmusic.*)',
        r'(?P<RockUUU0020music>.*\brock.*?\bmusic.*)',
        r'(?P<PopularUUU0020music>.*\bpopular.*?\bmusic.*)',
        r'(?P<IncidentalUUU0020music>.*\bincidental.*?\bmusic.*)',
        r'(?P<FolkUUU0020music>.*\bfolk.*?\bmusic.*)',
        r'(?P<DanceUUU0020music>.*\bdance.*?\bmusic.*)',
        r'(?P<ChamberUUU0020music>.*\bchamber.*?\bmusic.*)',
        r'(?P<ArtUUU0020music>.*\bart.*?\bmusic.*)',
        r'(?P<Music>.*\bmusic\b.*)',  # Different types of music are listed here
        r'(?P<MultivolumeUUU0020monograph>.*\bmulti[\-\s]*volume.*?\bmonograph.*)',
        r'(?P<MotionUUU0020picture>.*\b(motion.*?\bpicture|film|movie).*)',
        r'(?P<Model>.*\bmodel.*)',
        r'(?P<MilitaryUUU0020fiction>.*\bmilitary.*?\b(stor|fiction).*)',
        r'(?P<Military>.*\bmilitary.*)',
        r'(?P<MicroscopeUUU0020slide>.*\bmicroscope.*?\bslide.*)',
        r'(?P<MetaUUU0020fiction>.*\bmeta.*?\bfiction.*)',
        r'(?P<Memoir>.*\bmemoir.*)',
        r'(?P<Melodrama>.*\bmelo.*?\bdrama.*)',
        r'(?P<Map>.*\bmaps*\b.*)',
        r'(?P<Manuscript>.*\bmanuscript.*)',
        r'(?P<MagicalUUU0020realism>.*\bmagic.*?\brealism.*)',
        r'(?P<LooseUUU002Dleaf>.*\bloose[\-\s]*leaf.*)',
        r'(?P<Liturgy>.*\bliturg(y|ical).*)',
        r'(?P<Lithograph>.*\blitho[\-\s]*graph.*)',
        r'(?P<LifeUUU0020skillsUUU0020guide>.*\blife.*?\bskill.*?\bguide.*)',
        r'(?P<Libretto>.*\blibrett.*)',
        r'(?P<Letter>.*\bletter.*)',
        r'(?P<LegislativeUUU0020material>.*\blegislat.*?\bmaterial.*)',
        r'(?P<LegislativeUUU0020hearing>.*\blegislat.*?\bhearing.*)',
        r'(?P<LegislativeUUU0020address>.*\blegislat.*?\baddress.*)',
        r'(?P<Legislation>.*\blegislation.*)',
        r'(?P<LegalUUU0020digest>.*\b(legal|law).*?\bdigest.*)',
        r'(?P<LegalUUU0020caseUUU0020notes>.*\blegal.*?\bcase.*?\bnotes.*)',
        r'(?P<LegalUUU0020article>.*\blegal.*?\barticle.*)',
        r'(?P<Lecture>.*\blecture.*)',
        r'(?P<LearningUUU0020material>.*\blearn.*?\b(material|resource|guide).*)',
        r'(?P<LawUUU0020reportUUU0020orUUU0020digest>.*\blaw.*?\b(report|digest).*)',
        r'(?P<Law>^blaws*$)',   # Only matches whole string
        r'(?P<LargeUUU0020print>.*\blarge.*?\b(print|type).*)',
        r'(?P<LanguageUUU0020instruction>.*\blanguage.*?\binstruction.*)',
        r'(?P<Kit>.*\bkit.*)',
        r'(?P<Journal>.*\bjournal.*)',
        r'(?P<Issue>.*\bissue.*)',
        r'(?P<Interview>.*\binterview.*)',
        r'(?P<InstructionalUUU0020material>.*\binstruct.*?\b(material|resource|guide).*)',
        r'(?P<Instruction>.*\binstruction.*)',
        r'(?P<Index>.*\bindex.*)',
        r'(?P<Incunabulum>.*\bincunabul.*)',
        r'(?P<Illustration>.*\billustration.*)',
        r'(?P<Hymnal>.*\bhymnal.*)',
        r'(?P<HumourUUU0020orUUU0020satire>.*\b(humou?r|satir(e|ic)).*)',
        r'(?P<HorrorUUU0020fiction>.*\bhorror.*?\b(stor|fiction).*)',
        r'(?P<Horror>.*\bhorror.*)',
        r'(?P<HistoricalUUU0020fiction>.*\bhistor(ical|y)\s*fiction.*)',
        r'(?P<HistoricalUUU0020drama>.*\bhistori(ical|y).*?\bdrama.*)',
        r'(?P<History>.*\bhistory.*)',
        r'(?P<HighUUU0020interestUUU0020lowUUU0020vocabulary>.*\bhigh.*?\binterest.*?\blow.*?\bvocab.*)',
        r'(?P<HandbookUUU0020orUUU0020manual>.*\b(hand[\-\s]*book|manual).*)',
        r'(?P<Handbill>.*\bhand[\-\s]*bill.*)',
        r'(?P<Guidebook>.*\bguide[\-\s]*book.*)',
        r'(?P<Graphic>^graphic.*)',  # Only matches from start of string
        r'(?P<GovernmentUUU0020publicationUUU0020orUUU0020officalUUU0020document>.*\b(government.*?\bpublication|official.*?\bdocument).*)',
        r'(?P<GlossaryUUU0020orUUU0020vocabulary>.*\b(glossar|vocabular|terminolog|term.*?\bphrase).*)',
        r'(?P<Globe>.*\bglobe.*)',
        r'(?P<GhostUUU0020story>.*\bghost.*?\b(stor|fiction).*)',
        r'(?P<Gazetteer>.*\bgazetteer.*)',
        r'(?P<Gazette>.*\bgazettes*\b.*)',
        r'(?P<Gangster>.*\bgangster.*)',
        r'(?P<Game>.*\bgame.*)',
        r'(?P<Font>.*\bfont.*)',
        r'(?P<FolkUUU0020taleUUU0020orUUU0020fairyUUU0020tale>.*\b(fairy|folk)\s*(tale|lore).*)',
        r'(?P<Folio>.*\bfolio.*)',
        r'(?P<FlashUUU0020card>.*\bflash.*?\bcard.*)',
        r'(?P<FindingUUU0020aid>.*\bfinding.*?\baid.*)',
        r'(?P<Filmstrip>.*\bfilmstrip.*)',
        r'(?P<Filmography>.*\bfilmograph.*)',
        r'(?P<Fiction>^fiction.*)',  # Only matched from start of string
        r'(?P<Festschrift>.*\bfestschrift.*)',
        r'(?P<Fantasy>.*\bfantasy.*)',
        r'(?P<Facsimile>.*\bfacsim.*)',
        r'(?P<FableUUU0020orUUU0020legend>.*\b(fable|legend).*)',
        r'(?P<Exhibition>.*\bexhibi*tion.*)',
        r'(?P<Excerpt>.*\b(excerpt|extract|clips*\b).*)',
        r'(?P<ExaminationUUU0020questions>.*\bexam.*?\b(question|paper).*)',
        r'(?P<Eulogy>.*\beulog(y|ia|ies|ium).*)',
        r'(?P<Etching>.*\betching.*)',
        r'(?P<Essay>.*\bessay.*)',
        r'(?P<EroticUUU0020fiction>.*\berotic.*?\b(stor|fiction).*)',
        r'(?P<Erotic>.*\berotic\b.*)',
        r'(?P<Ephemera>.*\bephemera.*)',
        r'(?P<Engraving>.*\bengraving.*)',
        r'(?P<Encyclopaedia>.*\b(en)?c*yclopa*ed.*)',
        r'(?P<Ephemeris>.*\bephemeri(de)?s.*)',
        r'(?P<Elevation>.*\belevation.*)',
        r'(?P<DystopianUUU0020fiction>.*\bdystopia.*?\b(stor|fiction).*)',
        r'(?P<Dystopia>.*\bdystopia.*)',
        r'(?P<Drawing>.*\bdrawing.*)',
        r'(?P<Drama>.*\b(drama|plays*\b).*)',
        r'(?P<DomesticUUU0020fiction>.*\bdomestic.*?\b(stor|fiction).*)',
        r'(?P<DocuUUU002Ddrama>.*\bdocu[\-\s]*drama.*)',
        r'(?P<Documentary>.*\bdocumentar(y|ies).*)',
        r'(?P<DissertationUUU0020orUUU0020thesis>.*\b(dissertation|thes[ei]s).*)',
        r'(?P<Discography>.*\bdiscograph.*)',
        r'(?P<DisasterUUU0020fiction>.*\bdisaster.*?\b(stor|fiction).*)',
        r'(?P<Disaster>.*\bdisaster.*)',
        r'(?P<Directory>.*\bdir*ector(y|ies).*)',
        r'(?P<Diorama>.*\bdiorama.*)',
        r'(?P<Dictionary>.*\bdictionar.*)',
        r'(?P<DiaryUUU0020fiction>.*\bdiar(y|ies).*?\b(stor|fiction).*)',
        r'(?P<Diary>.*\bdiar(y|ies)\b.*)',
        r'(?P<DetectiveUUU0020andUUU0020mysteryUUU0020fiction>.*\bdetectiv.*?\bmyster.*?\b(stor|fiction).*)',
        r'(?P<DetectiveUUU0020andUUU0020mystery>.*\bdetectiv.*?\bmyster.*)',
        r'(?P<DetectiveUUU0020fiction>.*\bdetectiv.*?\b(stor|fiction).*)',
        r'(?P<Detective>.*\bdetectiv.*)',
        r'(?P<MysteryUUU0020fiction>.*\bmyster.*?\b(stor|fiction).*)',
        r'(?P<Mystery>.*\bmyster.*)',
        r'(?P<DerivativeUUU0020work>.*\bderivative.*?\bwork.*)',
        r'(?P<Deed>.*\bdeeds*\b.*)',
        r'(?P<Decision>.*\bdecision.*)',
        r'(?P<Debate>.*\bdebate.*)',
        r'(?P<DataUUU0020table>.*\b(data.*?\btable|table.*?\bdata).*)',
        r'(?P<Database>.*\bdata[\-\s]*base.*)',
        r'(?P<CriticsmUUU0020andUUU0020interpretation>.*\b(criticism|.*?\binterpretation).*)',
        r'(?P<CrimeUUU0020fiction>.*\bcrime.*?\b(stor|fiction).*)',
        r'(?P<Crime>.*\bcrime.*)',
        r'(?P<Correspondence>.*\bcorrespondence.*)',
        r'(?P<Cookery>.*\b(cook(ery|ing|book)?|recipe).*)',
        r'(?P<ConversionUUU0020table>.*\bconversion.*?\btable.*)',
        r'(?P<ConversionUUU0020chart>.*\bconversion.*?\bchart.*)',
        r'(?P<ConversationUUU0020orUUU0020phraseUUU0020book>.*\b(conversation|phrase).*?\bbook.*)',
        r'(?P<Congress>.*\bcongre*s+.*)',
        r'(?P<ConferenceUUU0020publication>.*\bconference.*?\b(publication|paper|proceeding|material|ephemer|transaction|program).*)',
        r'(?P<Concordance>.*\bconcordance.*)',
        r'(?P<CompendiumUUU0020orUUU0020compilation>.*\b(compend|collection|compilation).*)',
        r'(?P<CoparativeUUU0020study>.*\bcomparativ.*?\bstudy.*)',
        r'(?P<Commentary>.*\bcommentar.*)',
        r'(?P<ComicUUU0020orUUU0020graphicUUU0020novel>.*\b(comic|graphic.*?\bnovel).*)',
        r'(?P<Comedy>.*\bcomed(y|ies|ic)\b.*)',
        r'(?P<ColloquiumUUU0020publication>.*\bcolloqui.*?\b(publication|paper|proceeding|material|ephemer|transaction|program).*)',
        r'(?P<Codex>.*\b(codex|codices|codicil).*)',
        r'(?P<Chronology>.*\b(chronolog|time[\-\s]*line).*)',
        r'(?P<ChronicleUUU0020orUUU0020annal>.*\b(chronicle|annals*\b).*)',
        r'(?P<ChildrenUUU0027sUUU0020televisionUUU0020program>.*\b(children|juvenile).*?\bt(ele)?[.\s]*v(ision)?[.\s]*.*)',
        r'(?P<ChildrenUUU0027sUUU0020soundUUU0020recording>.*\b(children|juvenile).*?\bsound.*)',
        r'(?P<ChildrenUUU0027sUUU0020radioUUU0020program>.*\b(children|juvenile).*?\bradio.*)',
        r'(?P<ChildrenUUU0027sUUU0020map>.*\b(children|juvenile).*?\bmaps*\b.*)',
        r'(?P<ChildrenUUU0027sUUU0020literature>.*\b(children|juvenile).*?\bliterat*ure.*)',
        r'(?P<ChildrenUUU0027sUUU0020film>.*\b(children|juvenile).*?\b(motion.*?\bpicture|film|movie).*)',
        r'(?P<ChildrenUUU0027sUUU0020fiction>.*\b(children|juvenile).*?\b(stor|fiction).*)',
        r'(?P<ChildrenUUU0027sUUU0020drama>.*\b(children|juvenile).*?\b(drama|plays*\b).*)',
        r'(?P<ChildrenUUU0027sUUU0020book>.*\b(children|juvenile).*?\bbook.*)',
        r'(?P<ChildrenUUU0027sUUU0020audiobook>.*\b(children|juvenile).*?\baudio[\-\s]*book.*)',
        r'(?P<ChildrenUUU0027sUUU0020atlas>.*\b(children|juvenile).*?\batlas.*)',
        r'(?P<Chart>.*\bchart.*)',
        r'(?P<Chapbook>.*\bchap[\-\s]*book.*)',
        r'(?P<CensusUUU0020data>.*\bcensus.*)',
        r'(?P<CelestialUUU0020globe>.*\bcelestial.*?\bglobe.*)',
        r'(?P<CelestialUUU0020chart>.*\bcelestial.*?\bchart.*)',
        r'(?P<Catalogue>.*\bcatalog(ue)?.*)',
        r'(?P<CaseUUU0020study>.*\bcase.*?\b(study|report).*)',
        r'(?P<CartographicUUU0020material>.*\bcartographic.*?\bmaterial.*)',
        r'(?P<CaricatureUUU0020orUUU0020cartoon>.*\b(caricature|cartoon).*)',
        r'(?P<Calendar>.*\bcalendar.*)',
        r'(?P<ByUUU002Dlaw>.*\bbye*[\-\s]*law.*)',
        r'(?P<Broadside>.*\bbroadside.*)',
        r'(?P<BookUUU0020ofUUU0020hours>.*\bbook.*?\bhour.*)',
        r'(?P<BookUUU0020review>.*\bbook.*?\breview.*)',
        r'(?P<BookUUU0020list>.*\bbook.*?\blist.*)',
        r'(?P<BookUUU0020digest>.*\bbook.*?\bdigest.*)',
        r'(?P<BoardUUU0020book>.*\bboard.*?\bbook.*)',
        r'(?P<Blog>.*\bblog(s|ger|ging)?\b.*)',
        r'(?P<BirdUUU0027sUUU0020eyeUUU0020view>.*\bbird.*?\beye.*?\bview.*)',
        r'(?P<Biopic>.*\bbio[\-\s]*pic.*)',
        r'(?P<BiographicalUUU0020fiction>.*\bbo?iograph.*?\b(stor|fiction).*)',
        r'(?P<Biography>.*\bbo?iograph.*)',
        r'(?P<Bill>.*\bbills*\b.*)',
        r'(?P<Bildungsroman>.*\bbildungs[\-\s]*roman.*)',
        r'(?P<Bibliography>.*\bbibi?liograph.*)',
        r'(?P<Ballad>.*\bballad.*)',
        r'(?P<Autograph>.*\bauto[\-\s]*graph.*)',
        r'(?P<AutobiographicalUUU0020fiction>.*\bauto[\-\s]*biograph.*?\b(stor|fiction).*)',
        r'(?P<Autobiography>.*\bauto[\-\s]*biograph.*)',
        r'(?P<Aufsatzsammlung>.*\baufsatzsamml.*)',
        r'(?P<Audiobook>.*\baudio[\-\s]*book.*)',
        r'(?P<Atlas>.*\batlas.*)',
        r'(?P<ArtUUU0020reproduction>.*\bart.*?\breprod.*)',
        r'(?P<ArtUUU0020original>.*\bart.*?\boriginal.*)',
        r'(?P<Article>.*\barticle.*)',
        r'(?P<Aquatint>.*\baqua[\-\s]*tint.*)',
        r'(?P<Anthology>.*\bantholog.*)',
        r'(?P<Animation>.*\banimat(ion|ed).*)',
        r'(?P<Anecdote>.*\banecdote.*)',
        r'(?P<AlternativeUUU0020history>.*\bal*ternat.*?\bhistor.*)',
        r'(?P<Almanac>.*\balmanac.*)',
        r'(?P<AerialUUU0020view>.*\baerial.*?\bview.*)',
        r'(?P<AerialUUU0020photograph>.*\baerial.*?\bphotograph.*)',
        r'(?P<Advertisement>.*\b(advert|commercials).*)',
        r'(?P<Adventure>.*\badventure.*)',
        r'(?P<Address>.*\baddress.*)',
        r'(?P<Adaptation>.*\badapt(at)?ion.*)',
        r'(?P<ActivityUUU0020program>.*\bactivity.*?\bprogram.*)',
        r'(?P<AbstractUUU0020orUUU0020summary>.*\b(abstract|summary).*)',
        r'(?P<AllElse>.*)',
        )

    def TranslationsInto(self, mo):
        lang = re.sub(r'.*\btranslations.*?\binto.*?\b([a-z]+).*', r'\1', mo.group(), flags=re.IGNORECASE)
        lang = lang[0].upper() + lang[1:].lower()
        return 'Translations into {}'.format(lang)

    def TranslationsFrom(self, mo):
        lang = re.sub(r'.*\btranslations.*?\bfrom.*?\b([a-z]+).*', r'\1', mo.group(), flags=re.IGNORECASE)
        lang = lang[0].upper() + lang[1:].lower()
        return 'Translations from {}'.format(lang)

    def EarlyWorks(self, mo):
        return 'Early works to ' + re.sub(r'.*\bearly.*?\bworks.*?\bto.*?\b([0-9]+).*', r'\1', mo.group(), flags=re.IGNORECASE)


class PlaceNamesUK(MultiRegex):    
    regexes = (
        r'(?P<County>,* \(?Co\.? (Antrim|Armagh|Carlow|Cavan|Clare|Cork|Donegal|Down|Dublin|Fermanagh|Galway|Kerry|Kildare|Kilkenny|Laois|Laoighis|Leitrim|Limerick|Londonderry|Longford|Louth|Mayo|Meath|Monaghan|Offaly|Roscommon|Sligo|Tipperary|Tyrone|Waterford|Westmeath|Wexford|Wicklow)\)?\b\s*([.,]|$))',
        r'(?P<CountyB>Co\.? (Antrim|Armagh|Carlow|Cavan|Clare|Cork|Donegal|Down|Dublin|Fermanagh|Galway|Kerry|Kildare|Kilkenny|Laois|Laoighis|Leitrim|Limerick|Londonderry|Longford|Louth|Mayo|Meath|Monaghan|Offaly|Roscommon|Sligo|Tipperary|Tyrone|Waterford|Westmeath|Wexford|Wicklow)\b)',
        r'(?P<UUU002CUUU0020Bedfordshire>,* \(?beds\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Berkshire>,* \(?berks\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Buckinghamshire>,* \(?bucks\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Cambridgeshire>,* \(?cambs\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020CountyUUU0020Durham>,* \(?co\.*\s*durham\)?\b\.*)'
        r'(?P<UUU002CUUU0020EastUUU0020Sussex>,* \(?e\.*\s*sussex\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Gloucestershire>,* \(?glos\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Hampshire>,* \(?hants\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Hertfordshire>,* \(?herts\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Lancashire>,* \(?lancs\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Leicestershire>,* \(?leics\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Lincolnhire>,* \(?lincs\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Middlesex>,* \(?middx?\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NorthUUU0020Humberside>,* \(?n\.*\s*humbs\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NorthUUU0020Yorkshire>,* \(?n\.*\s*yorks\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Northamptonshire>,* \(?northants\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Northumberland>,* \(?northd\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Nottinghamshire>,* \(?notts\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Oxfordshire>,* \(?oxon\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Shropshire>,* \(?salop\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020SouthHumberside>,* \(?s\.*\s*humbs\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020SouthUUU0020Yorkshire>,* \(?s\.*\s*yorks\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Staffordshire>,* \(?staffs\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Warwickshire>,* \(?warks\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020WestUUU0020Midlands>,* \(?w\.*\s*mids\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020WestUUU0020Sussex>,* \(?w\.*\s*wussex\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020WestUUU0020Yorkshire>,* \(?w\.*\s*yorks\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Wiltshire>,* \(?wilts\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Worcestershire>,* \(?worcs\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020MidUUU0020Glamorgan>,* \(?m\.*\s*glam\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020SouthUUU0020Glamorgan>,* \(?s\.*\s*glam\)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020WestUUU0020Glamorgan>,* \(?w\.*\s*glam\)?\b\s*([.,]|$))',
        r'(?P<BerwickUUU002DuponUUU002DTweed>berwick[\-\s]*upon[\-\s]*tweed)',
        r'(?P<Blairgowrie>blairgowe?rie)',
        r'(?P<BuryUUU0020StUUU0020Edmunds>Bury St\.* Edmund[\.\'s]*)',
        r'(?P<Haverfordwest>haverford\s*west)',
        r'(?P<HytheUUU002CUUU0020Southampton>hythe,?\s*southampton)',
        r'(?P<KingUUU0027sUUU0020Lynn>[kK]ing\'?s [lL]ynn)',
        r'(?P<London>\b(londres\b|londra\b|london,?\s*(u\.*k\.*|england)))',
        r'(?P<LythamUUU0020StUUU0020AnneUUU0027s>Lytham St\.* Anne[\.\'s]*)',
        r'(?P<NewcastleUUU0020uponUUU0020Tyne>\bnewcastle[\-\s]*(up)?on[\-\s]*tyne\b)',
        r'(?P<UUU0020Scotland>,*\s+s[.\s]*w[.\s]*scotland)',
        r'(?P<SaffronUUU0020Walden>saffron\s*wald[oe]n)',
        r'(?P<StratfordUUU0020uponUUU0020Avon>\bstrat?ford[\-\s]*(up)?on[\-\s]*avon\b)',
        r'(?P<StokeUUU002DonUUU002DTrent>\bstoke[\-\s]*(up)?on[\-\s]*trent\b)',
        r'(?P<StocktonUUU002DonUUU002DTees>stockton[\-\s]*on[\-\s]*tees)',
        r'(?P<TunbridgeUUU0020Wells>tunbridge\s*wells)',
        r'(?P<WhitleyUUU0020Bay>whit(el|le)y\s*bay)',
        r'(?P<UUU0020LaneUUU002CUUU0020>\s+l[an]\.?,\s+)',
        r'(?P<UUU0020PlaceUUU002CUUU0020>\s+pl\.?,\s+)',
        r'(?P<UUU0020RoadUUU002CUUU0020>\s+rd?\.?,\s+)',
        r'(?P<UUU0020StreetUUU002CUUU0020>\s+st\.?,\s+)',
        r'(?P<UUU0020SquareUUU002CUUU0020>\s+sq\.?,\s+)',
        r'(?P<POUUU0020Box>P\.?O\.?\s*box)',
    )

    def County(self, mo):
        county = re.sub(r',* \(?Co\.? (Antrim|Armagh|Carlow|Cavan|Clare|Cork|Donegal|Down|Dublin|Fermanagh|Galway|Kerry|Kildare|Kilkenny|Laois|Laoighis|Leitrim|Limerick|Londonderry|Longford|Louth|Mayo|Meath|Monaghan|Offaly|Roscommon|Sligo|Tipperary|Tyrone|Waterford|Westmeath|Wexford|Wicklow)\)?\b\.*',
                        r'\1', mo.group(), flags=re.IGNORECASE)
        county = county[0].upper() + county[1:].lower()
        return ' County {}'.format(county)
    
    def CountyB(self, mo):
        try:
            county = re.sub(r'Co\.? (Antrim|Armagh|Carlow|Cavan|Clare|Cork|Donegal|Down|Dublin|Fermanagh|Galway|Kerry|Kildare|Kilkenny|Laois|Laoighis|Leitrim|Limerick|Londonderry|Longford|Louth|Mayo|Meath|Monaghan|Offaly|Roscommon|Sligo|Tipperary|Tyrone|Waterford|Westmeath|Wexford|Wicklow)\b)',
                            r'\1', mo.group(), flags=re.IGNORECASE)
        except: return ''
        county = county[0].upper() + county[1:].lower()
        return ' County {}'.format(county)


class PlaceNamesUS(MultiRegex):
    regexes = (
        r'(?P<UUU002CUUU0020Alabama>[,\s]*(?<!-)\b(ala|a[.\s]*l)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Alaska>[,\s]*(?<!-)\ba[.\s]*k\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Arizona>[,\s]*(?<!-)\b(ariz|a[.\s]*z)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Arkansas>[,\s]*(?<!-)\b(ark|a[.\s]*r)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Bermuda>[,\s]*(?<!-)\bb[.\s]*e\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020California>[,\s]*(?<!-)\b(calif|claif|c[.\s]*a)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Colorado>[,\s]*(?<!-)\b(colo|c[.\s]*o)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Connecticut>[,\s]*(?<!-)\b(conn|c[.\s]*t)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Delaware>[,\s]*(?<!-)\b(del|d[.\s]*e)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020DistrictUUU0020ofUUU0020Columbia>[,\s]*(?<!-)\bd[.\s]*c\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Florida>[,\s]*(?<!-)\b(fla|f[.\s]*l)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Georgia>[,\s]*(?<!-)\bg[.\s]*a\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Hawaii>[,\s]*(?<!-)\bh[.\s]*i\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Idaho>[,\s]*(?<!-)\bi[.\s]*d\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Illinois>[,\s]*(?<!-)\b(ill|i[.\s]*l)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Indiana>[,\s]*(?<!-)\b(ind|i[.\s]*n)\b(\.|[.\s]*$))',
        r'(?P<UUU002CUUU0020Iowa>[,\s]*(?<!-)\bi[.\s]*a\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Kansas>[,\s]*(?<!-)\b(kan|k[.\s]*s)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Kentucky>[,\s]*(?<!-)\bk[.\s]*y\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Louisiana>[,\s]*(?<!-)\bl[.\s]*a\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Maine>[,\s]*(?<!-)\bm[.\s]*e\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Maryland>[,\s]*(?<!-)\bm[.\s]*d\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Massachusetts>[,\s]*(?<!-)\b(mass|m[.\s]*a)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Michigan>[,\s]*(?<!-)\b(mich|m[.\s]*i)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Minnesota>[,\s]*(?<!-)\b(minn|m[.\s]*n)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Mississippi>[,\s]*(?<!-)\b(miss|m[.\s]*s)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Missouri>[,\s]*(?<!-)\bm[.\s]*o\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Montana>[,\s]*(?<!-)\b(mont|m[.\s]*t)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Nebraska>[,\s]*(?<!-)\b(neb|n[.\s]*e)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Nevada>[,\s]*(?<!-)\b(nev|n[.\s]*v)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NewUUU0020Hampshire>[,\s]*(?<!-)\bn[.\s]*h\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NewUUU0020Jersey>[,\s]*(?<!-)\bn[.\s]*j\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NewUUU0020Mexico>[,\s]*(?<!-)\bn[.\s]*m\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NewUUU0020York>[,\s]*(?<!-)\bn[.\s]*y\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NorthUUU0020Carolina>[,\s]*(?<!-)\bn[.\s]*c\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NorthUUU0020Dakota>[,\s]*(?<!-)\bn[.\s]*d\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Ohio>[,\s]*(?<!-)\bo[.\s]*h\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Oklahoma>[,\s]*(?<!-)\b(okla|o[.\s]*k)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Oregon>[,\s]*(?<!-)\bo[.\s]*r\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Pennsylvania>[,\s]*(?<!-)\bp[.\s]*a\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020PuertoUUU0020Rico>[,\s]*(?<!-)\bp[.\s]*r\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020RhodeUUU0020Island>[,\s]*(?<!-)\br[.\s]*i\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020SouthUUU0020Carolina>[,\s]*(?<!-)\bs[.\s]*c\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020SouthUUU0020Dakota>[,\s]*(?<!-)\bs[.\s]*d\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Tennessee>[,\s]*(?<!(-| by ))\b(tenn|t[.\s]*n)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Texas>[,\s]*(?<!-)\b(tex|t[.\s]*x)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Utah>[,\s]*(?<!-)\bu[.\s]*t\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Vermont>[,\s]*(?<!-)\bv[.\s]*t\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020VirginUUU0020Island>[,\s]*(?<!-)\bv[.\s]*i\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Virginia>[,\s]*(?<!-)\bv[.\s]*a\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Washington>[,\s]*(?<!-)\b(wash|w[.\s]*a)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020WestUUU0020Virginia>[,\s]*(?<!-)\bw[.\s]*va?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Wisconsin>[,\s]*(?<!-)\b(wis|w[.\s]*i)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Wyoming>[,\s]*(?<!-)\b(wyo|w[.\s]*y)\b\s*([.,]|$))',        
    )


class PlaceNamesAustralia(MultiRegex):
    regexes = (
        r'(?P<UUU002CUUU0020NewUUU0020SouthUUU0020Wales>[,\s]*(?<!-)\bn[.\s]*s[.\s]*w\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Queensland>[,\s]*(?<!-)\bq[.\s]*l[.\s]*d\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020SouthUUU0020Australia>[,\s]*(?<!-)\bs[.\s]*a(ust)?\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Tasmania>[,\s]*(?<!-)\bt[.\s]*a[.\s]*s\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Victoria>[,\s]*(?<!-)\bv[.\s]*i[.\s]*c\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020WesternUUU0020Australia>[,\s]*(?<!-)\bw[.\s]*a\b\s*([.,]|$))',
    )


class PlaceNamesBrazil(MultiRegex):
    regexes = (
        r'(?P<UUU002CUUU0020Acre>[,\s]*(?<!-)\ba[.\s]*c\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Alagoas>[,\s]*(?<!-)\ba[.\s]*l\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020AmapUUU00E1>[,\s]*(?<!-)\ba[.\s]*p\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Amazonas>[,\s]*(?<!-)\ba[.\s]*m\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Bahia>[,\s]*(?<!-)\bb[.\s]*a\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020CearUUU00E1>[,\s]*(?<!-)\bc[.\s]*e\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020DistritoUUU0020Federal>[,\s]*(?<!-)\bd[.\s]*f\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020EspUUU00EDritoUUU0020Santo>[,\s]*(?<!-)\be[.\s]*s\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020GoiUUU00E1s>[,\s]*(?<!-)\bg[.\s]*o\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020MaranhUUU00E3o>[,\s]*(?<!-)\bm[.\s]*a\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020MatoUUU0020Grosso>[,\s]*(?<!-)\bm[.\s]*t\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020MatoUUU0020GrossoUUU0020doUUU0020Sul>[,\s]*(?<!-)\bm[.\s]*s\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020MinasUUU0020Gerais>[,\s]*(?<!-)\bm[.\s]*g\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020ParUUU00E1>[,\s]*(?<!-)\bp[.\s]*a\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020ParaUUU00EDba>[,\s]*(?<!-)\bp[.\s]*b\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020ParanUUU00E1>[,\s]*(?<!-)\bp[.\s]*r\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Pernambuco>[,\s]*(?<!-)\bp[.\s]*e\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020PiauUUU00ED>[,\s]*(?<!-)\bp[.\s]*i\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020RioUUU0020deUUU0020Janeiro>[,\s]*(?<!-)\br[.\s]*j\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020RioUUU0020GrandeUUU0020doUUU0020Norte>[,\s]*(?<!-)\br[.\s]*n\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020RioUUU0020GrandeUUU0020doUUU0020Sul>[,\s]*(?<!-)\br[.\s]*s\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020RondUUU00F4nia>[,\s]*(?<!-)\br[.\s]*o\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Roraima>[,\s]*(?<!-)\br[.\s]*r\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020SantaUUU0020Catarina>[,\s]*(?<!-)\bs[.\s]*c\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020SUUU00E3oUUU0020Paulo>[,\s]*(?<!-)\bs[.\s]*p\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Sergipe>[,\s]*(?<!-)\bs[.\s]*e\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Tocantins>[,\s]*(?<!-)\bt[.\s]*o\b\s*([.,]|$))',
    )    


class PlaceNamesCanada(MultiRegex):
    regexes = (
        r'(?P<UUU002CUUU0020Alberta>[,\s]*(?<!-)\b(a(lt|tl)a|a[.\s]*b)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020BritishUUU0020Columbia>[,\s]*(?<!-)\bb[.\s]*c\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Manitoba>[,\s]*(?<!-)\b(man|m[.\s]*b)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NewUUU0020Brunswick>[,\s]*(?<!-)\bn[.\s]*b\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NewfoundlandUUU0020andUUU0020Labrador>[,\s]*(?<!-)\bn[.\s]*l\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NovaUUU0020Scotia>[,\s]*(?<!-)\bn[.\s]*s\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020NorthwestUUU0020Territories>[,\s]*(?<!-)\bn[.\s]*(w[.\s]*)?t\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Nunavut>[,\s]*(?<!-)\bn[.\s]*u\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Ontario>[,\s]*(?<!-)\bo[.\s]*n[.\s]*(t\b)?\.*)',
        r'(?P<UUU002CUUU0020PrinceUUU0020EdwardUUU0020Island>[,\s]*(?<!-)\bp[.\s]*e\b[.\s]*(i\b)?\.*)',
        r'(?P<UUU002CUUU0020Quebec>[,\s]*(?<!-)\b(que|q[.\s]*c)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Saskatchewan>[,\s]*(?<!-)\b(sask|s[.\s]*k)\b\s*([.,]|$))',
        r'(?P<UUU002CUUU0020Yukon>[,\s]*(?<!-)\by[.\s]*t\b\s*([.,]|$))',
    )    


class PlaceNamesNewZealand(MultiRegex):
    regexes = (
        r'(?P<UUU002CUUU0020NewUUU0020Zealand>[,\s]*(?<!-)\bn[.\s]*z\b\s*([.,]|$))',
    )


class PlaceNamesOther(MultiRegex):
    regexes = (
        # Russia
        r'(?P<SaintUUU0020Petersburg>(\bs(ai?nk?)?t?\.?-?\s*peters?burg.*|^spb$))',
        # Other
        r'(?P<BadUUU0020HomburgUUU0020vorUUU0020derUUU0020HUUU00F6he>\bbad\s*homburg\s*v(or)?[.\s]*d(er)?[.\s]*h(.he)?\b[.]*)',
        r'(?P<CambridgeUUU002CUUU0020Massachusetts>\bcambridge[,\s]+(mass|m[.\s]*a)\b\s*([.,]|$))',
        r'(?P<FrankfurtUUU0020amUUU0020Main>\bfran[ck]f[uo]rt\s*[ao]m*\.*\s*m(ain)?\b\.*)',
        r'(?P<FreiburgUUU0020imUUU0020Breisgau>\bfreib(urg)?\.*\s*i[mn]*\.*\s*br(eisgau)?\b\.*)',
        r'(?P<UUU002CUUU0020NewUUU0020Zealand>[,\s]*(?<!-)\bn[.\s]*z\b\s*([.,]|$))',
        r'(?P<SantiagoUUU002CUUU0020DominicanUUU0020Republic>\bsantiago,? rd\b)',
        r'(?P<TUUU00FCbingen>\bt.e?bi+ngen\b)',
        r'(?P<TUUU02bfbilisi>\bt.*?bilisi\b)',
    )


class PlaceNamesAccents(MultiRegex):
    regexes = (
        r'(?P<UUU00C5enUUU00E5>A\.benra\.)',
        r'(?P<AbUUU016bUUU0020UUU1e92aby>Abu\. Z\.aby)',
        r'(?P<UUU00C5lborg>A\.lborg)',
        r'(?P<AngoulUUU00EAme>Angoule\.me)',
        r'(?P<Antsiranana>Antsiran\.ana)',
        r'(?P<UUU00C5rhus>A\.rhus)',
        r'(?P<ArkhangelUUU02B9sk>Arkh?angel\.sk)',
        r'(?P<AsunciUUU00F3n>Asuncio\.n)',
        r'(?P<BaUUU1E63rah>Bas\.rah)',
        r'(?P<BUUU00E9kUUU00E9scsaba>Be\.ke\.scsaba)',
        r'(?P<BesanUUU00E7on>Besanc\.on)',
        r'(?P<BiaUUU0142ystok>Bia\.ystok)',
        r'(?P<BolesUUU0142awiec>Boles\.awiec)',
        r'(?P<BrasUUU00EDlia>Brasi\.lia)',
        r'(?P<BraUUU0219ov>Bras\.ov)',
        r'(?P<BremervUUU00F6rde>Bremervo\.rde)',
        r'(?P<brUUU00FCck>bru\.ck)',
        r'(?P<BystrzycaUUU0020Kodzka>Bystrzyca K\.odzka)',
        r'(?P<CUUU00E1diz>Ca\.diz)',
        r'(?P<CaUUU00F1ete>Can\.ete)',
        r'(?P<CUUU00E9ret>Ce\.ret)',
        r'(?P<ChambUUU00E9ry>Chambe\.ry)',
        r'(?P<ChernivUUU0074UUU0073UUU0069>Chernivt\.s\.i)',
        r'(?P<ChiUUU0219inUUU0103u>Chis\.ina\.u)',
        r'(?P<ChojnUUU00F3w>Chojno\.w)',
        r'(?P<ConnahUUU0027sUUU0020Quay>Connah\.s Quay)',
        r'(?P<CUUU00f4teUUU0020dUUU0027Ivoire>Co\.te d\.Ivoire)',
        r'(?P<CUUU00FAcuta>Cu.cuta)',
        r'(?P<CuraUUU00E7ao>Curac\.ao)',
        r'(?P<DawUUU1E25ah>Dawh\.ah)',
        r'(?P<DUUU00F6beln>Do\.beln)',
        r'(?P<DolnoUUU015BlUUU0105skie>Dolnos\.la\.skie)',
        r'(?P<DonauwUUU00F6rth>Donauwo\.rth)',
        r'(?P<DUUU00FAnUUU0020Laoghaire>Du\.n Laoghaire)',
        r'(?P<DUUU00FCsseldorf>Du\.sseldorf)',
        r'(?P<DzierUUU017ConiUUU00F3w>Dzierz\.onio\.w)',
        r'(?P<ElblUUU0105g>Elbla\.g)',
        r'(?P<Eleia>E\.leia)',
        r'(?P<EntreUUU0020RUUU00edos>Entre Ri\.os)',
        r'(?P<FUUU00FCrstenfeld>Fu\.rstenfeld)',
        r'(?P<FUUU00FCrstenwalde>Fu\.rstenwalde)',
        r'(?P<FUUU00FCrth>Fu\.rth)',
        r'(?P<GdaUUU0144sk>Gdan\.sk)',
        r'(?P<GmUUU00FCnd>Gmu\.nd)',
        r'(?P<GUUU0142ogUUU00F3w>G\.ogo\.w)',
        r'(?P<GUUU00F6ppingen>Go\.ppingen)',
        r'(?P<GUUU00F6rlitz>Go\.rlitz)',
        r'(?P<GUUU00F6teborg>Go\.teborg)',
        r'(?P<GUUU00F6ttingen>Go\.ttingen)',
        r'(?P<GrUUU00F6bming>Gro\.bming)',
        r'(?P<GUUU00FCstrow>Gu\.strow)',
        r'(?P<GUUU00F6tersloh>Gu\.tersloh)',
        r'(?P<GyUUU0151r>Gyo\.r)',
        r'(?P<UUU1E24alfUUU0101UUU02BCalUUU002DJadUUU012Bdah>H\.alfa\.\. al\-Jadi\.dah)',
        r'(?P<UUU0126amrun>H\.amrun)',
        r'(?P<HUUU00F3dmezUUU0151vUUU00E1sUUU00E1rhely>Ho\.dmezo\.va\.sa\.rhely)',
        r'(?P<IAroslavlUUU0027>I\.A\.roslavl\.)',
        r'(?P<IAroslavskaia>I\.A\.roslavskai\.a\.)',
        r'(?P<UUU00CDsafjUUU00F6rUUU00F0ur>I\.safjo\.rur)',
        r'(?P<UUU0130zmir>I\.zmir)',
        r'(?P<JesenUUU00EDk>Jeseni\.k)',
        r'(?P<JUUU00FClich>Ju\.lich)',
        r'(?P<KlaipUUU0117da>Klaipe\.da)',
        r'(?P<KUUU014DbeUUU002Dshi>Ko\.be\-shi)',
        r'(?P<KUUU00F6nigs>Ko\.nigs)',
        r'(?P<KoUUU0142obrzeg>Ko\.obrzeg)',
        r'(?P<KoUUU0161ice>Kos\.ice)',
        r'(?P<KUUU00F6then>Ko\.then)',
        r'(?P<Krasnoiarsk>Krasnoi\.a\.rsk)',
        r'(?P<LaUUU0020CoruUUU00F1a>La Corun\.a)',
        r'(?P<LeUUU015Bna>Les\.na)',
        r'(?P<LiUUU00E8ge>Lie\.ge)',
        r'(?P<LogroUUU00F1o>Logron\.o)',
        r'(?P<LomUUU00E9>Lome\.)',
        r'(?P<LUUU00FCbben>Lu\.bben)',
        r'(?P<LUUU00FCbeck>Lu\.beck)',
        r'(?P<LUUU00FCdenscheid>Lu\.denscheid)',
        r'(?P<LUUU00FCneburg>Lu\.neburg)',
        r'(?P<LUUU00FCtzen>Lu\.tzen)',
        r'(?P<Lviv>L\.viv)',
        r'(?P<MUUU00E1laga>Ma\.laga)',
        r'(?P<MarkranstUUU00E4dt>Markransta\.dt)',
        r'(?P<MUUU00E9ziUUU00E8res>Me\.zie\.res)',
        r'(?P<MieroszUUU00F3w>Mieroszo\.w)',
        r'(?P<MontbUUU00E9liard>Montbe\.liard)',
        r'(?P<MontluUUU00E7on>Montluc\.on)',
        r'(?P<MontrUUU00E9al>Montre\.al)',
        r'(?P<MoUUU00FBtiers>Mou\.tiers)',
        r'(?P<MUUU00FClheim>Mu\.lheim)',
        r'(?P<MUUU00FCnden>Mu\.nden)',        
        r'(?P<MUUU00FCnster>Mu\.nster)',
        r'(?P<NeuchUUU00E2tel>\bneucha\.tel\b)',
        r'(?P<NUUU00EEmes>\bni\.mes\b)',
        r'(?P<NottinghamshireUUU0020UUU0028EnglandUUU0029>Nottinghamshire, Eng\b)',
        r'(?P<NoumUUU00E9a>\bnoume\.a\b)',
        r'(?P<NukuUUU0027alofa>Nuku\.alofa)',
        r'(?P<NUUU00FCrtingen>\bnu\.rtingen\b)',
        r'(?P<oblastUUU02B9>oblast\.)',
        r'(?P<OdrzaUUU0144skie>\bodrzan\.skie\b)',
        r'(?P<UUU0141UUU00F3dUUU017a>\bo\.dz\.)',
        r'(?P<PaUUU00EDsUUU0020Vasco>\bpai\.s\s*vasco\b)',
        r'(?P<ParaUUU00EDba>Parai\.ba)',
        r'(?P<PUUU00E9rigueux>Pe\.rigueux)',
        r'(?P<PUUU00E9ruwelz>Pe\.ruwelz)',
        r'(?P<PiotrkUUU00F3w>Piotrko\.w)',
        r'(?P<PUUU0142ock>P\.ock)',
        r'(?P<PortoUUU0020UniUUU00E3o>Porto Unia\.o)',
        r'(?P<PUUU00F6ssneck>Po\.ssneck)',
        r'(?P<PUUU0027yUUU014Fngyang>P\.yo\.ngyang)',
        r'(?P<QuUUU0027Appelle>Qu\.Appelle)',
        r'(?P<QuUUU00E9bec>Que\.bec)',
        r'(?P<QuerUUU00E9taro>Quere\.taro)',
        r'(?P<ReUUU0219iUUU021Ba>Res\.it\.a)',
        r'(?P<RUUU00E9union>Re\.union)',
        r'(?P<ReykjavUUU00EDk>Reykjavi\.k)',
        r'(?P<RibeirUUU00E3oUUU0020Preto>Ribeira\.o Pre\.to)',
        r'(?P<RUUU012Bga>Ri\.ga)',
        r'(?P<SanUUU0020JuanUUU0020delUUU0020Rio>San Juan del Rio\.)',
        r'(?P<SanktUUU0020PUUU00F6lten>Sankt Po\.lten)',
        r'(?P<SanUUU0020SebastiUUU00E1n>San Sebastia\.n)',
        r'(?P<SUUU00E3oUUU0020JoUUU00E3oUUU0020delUUU0020Rei>Sa\.o Joa\.o del Rei)',
        r'(?P<SUUU00E3oUUU0020Paulo>Sa\.o Paulo)',
        r'(?P<SchUUU00F6nebeck>Scho\.nebeck)',
        r'(?P<ShenkurskiiUUU0020raion>Shenkurskii\. rai\.on)',
        r'(?P<SUUU015BlUUU0105skie>S\.la\.skie)',
        r'(?P<TUUU02BFbilisi>T\.bilisi)',
        r'(?P<TUUU00E9touan>Te\.touan)',
        r'(?P<TimiUUU0219oara>Timis\.oara)',
        r'(?P<TUUU00EErguUUU002DMureUUU0219>Ti\.rgu\-Mures\.)',
        r'(?P<Tiruchchirappalli>Tiruchchira\.ppalli)',
        r'(?P<TUUU00F3rshavn>To\.rshavn)',
        r'(?P<TUUU00FCbingen>Tu\.bingen)',
        r'(?P<UzUUU00E8s>Uze\.s)',
        r'(?P<Uzhhorod>Uz\.h\.horod)',
        r'(?P<ValparaUUU00EDso>Valparai\.so)',
        r'(?P<VeszprUUU00E9m>Veszpre\.m)',
        r'(?P<WaUUU0142brzych>Wa\.brzych)',
        r'(?P<WaUUU0142cz>Wa\.cz)',
        r'(?P<WarmiUUU0144sko>Warmin\.sko)',
        r'(?P<WojewUUU00F3dztwo>Wojewo\.dztwo)',
        r'(?P<WrocUUU0142aw>Wroc\.aw)',
        r'(?P<WUUU00FCrzburg>Wu\.rzburg)',
        r'(?P<ZUUU017BagaUUU0144>Z\.agan\.)',
        r'(?P<ZiUUU0119bice>Zie\.bice)',
        r'(?P<Zielonauuu0020GUUU00F3ra>Zielona Go\.ra)',
    )


class Relators(MultiRegex):
    simple = True
    regexes = (
        r'(?P<authorUUU0020ofUUU0020introduction>^intro(duc)?(ed|tion)?\.*$)',
        r'(?P<adaptor>.*adapt[eo]r.*)',
        r'(?P<artist>.*artist.*)',
        r'(?P<attributedUUU0020name>^atr$)',
        r'(?P<author>.*auth?(o|eu)r.*)',
        r'(?P<bibliographer>.*bibliographer.*)',
        r'(?P<biographer>.*biographer.*)',
        r'(?P<bookseller>.*bookseller.*)',
        r'(?P<cartographer>.*cartographer.*)',
        r'(?P<cartoonist>.*cartoonist.*)',
        r'(?P<crosswordUUU0020compiler>.*crossword[\-\s]*(puzzle)?[\-\s]*compiler.*)',
        r'(?P<compiler>(.*compiler.*|^(joint\s*)?comp?s?$))',
        r'(?P<composer>.*composer.*)',
        r'(?P<correspondent>.*correspond[ea]nt.*)',
        r'(?P<creator>.*creator.*)',
        r'(?P<critic>.*critic.*)',
        r'(?P<draughtsman>.*draughtsman.*)',
        r'(?P<assistantUUU0020editor>.*assistant[\-\s]*editor.*)',
        r'(?P<associateUUU0020editor>.*associate[\-\s]*editor.*)',
        r'(?P<editor>(.*(editor|directeur de la publication|diteur).*|^(ed|ditor)$))',
        r'(?P<engineer>.*engineer.*)',
        r'(?P<engraver>(.*engraver.*|^engr$))',
        r'(?P<historian>.*historian.*)',
        r'(?P<illustrator>.*illustrator.*)',
        r'(?P<interviewee>.*interviewee.*)',
        r'(?P<interviewer>.*interviewer.*)',
        r'(?P<journalist>.*journalist.*)',
        r'(?P<lecturer>.*lecturer.*)',
        r'(?P<musician>.*musician.*)',
        r'(?P<narrator>.*narrator.*)',
        r'(?P<novelist>.*novelist.*)',
        r'(?P<painter>.*painter.*)',
        r'(?P<performer>.*performer.*)',
        r'(?P<photographer>.*photographer.*)',
        r'(?P<playwright>.*playwright.*)',
        r'(?P<poet>.*poet.*)',
        r'(?P<printer>(.*printer.*|^pt*r$))',
        r'(?P<printmaker>.*printmaker.*)',
        r'(?P<publisher>.*publisher.*)',
        r'(?P<screenwriter>.*screen-*writer.*)',
        r'(?P<surveyor>.*surveyor.*)',
        r'(?P<teacher>.*teacher.*)',
        r'(?P<topographer>.*topographer.*)',
        r'(?P<translator>(.*translator.*|^tr(ans|aduction)$))',
        r'(?P<typesetter>.*typesetter.*)',
        r'(?P<writerOf>.*writer\s*of\s*(.*?)\s*(,|and).*)',
        r'(?P<writerOn>.*writer\s*on\s*(.*?)\s*(,|and).*)',
        r'(?P<writer>.*writer.*)',
        r'(?P<AllElse>.*)',
    )

    def writerOf(self, mo):
        text = re.sub(r'.*writer\s*of\s*(.*?)\s*(,|and).*', r'\1', mo.group(), flags=re.IGNORECASE)
        return 'writer of {}'.format(text)

    def writerOn(self, mo):
        text = re.sub(r'.*writer\s*on\s*(.*?)\s*(,|and).*', r'\1', mo.group(), flags=re.IGNORECASE)
        return 'writer of {}'.format(text)
