import kivy
kivy.require("1.10.1")

from kivy.app import App
from kivy.uix.pagelayout import PageLayout
from kivy.properties import NumericProperty, ObjectProperty, ListProperty, StringProperty, BooleanProperty
from kivy.clock import Clock
from kivy.uix.label import Label
from kivy.config import Config

import datetime
import winsound
from yr.libyr import Yr
import ghostscript
import locale

import pygame

from pyo365 import Account, Connection

from bs4 import BeautifulSoup
import re

from credentials import credentails
scopes = ['https://graph.microsoft.com/Mail.ReadWrite', "offline_access"]

account = Account(credentials, scopes=scopes)
con = Connection(credentials, scopes=scopes)

class PagesClass(PageLayout):
    time = ObjectProperty(datetime.datetime.now(), True)
    waited = ObjectProperty(datetime.datetime.now())

    alarmHour = ObjectProperty(None)
    alarmMinute = ObjectProperty(None)
    alarmText = ObjectProperty(None)
    alarm = BooleanProperty(False)

    currentDay = NumericProperty(7)
    schedulePlan = ObjectProperty(None)
    schedulePlans = ListProperty([["Musikk", "Naturfag", "Engelsk"], 
                                  ["Matematikk", "Matematikk", "Norsk", "Valgfag", "Valgfag"], 
                                  ["Matematikk", "Fremmedspråk", "Naturfag", "Norsk", "Norsk"], 
                                  ["Engelsk", "Fremmedspråk", "Samfunnsfag", "Kroppsøving", "Kroppsøving"], 
                                  ["Kunst Og Håndverk", "Kunst Og Håndverk", "Utdanningsvalg", "Samfunnsfag"]])

    currentHour = NumericProperty(None)
    weatherSymbol = StringProperty("1")
    degrees = NumericProperty(0)

    days = ListProperty(["Mandag", "Tirsdag", "Onsdag", "Torsdag", "Fredag", "Lørdag", "Søndag", ""])
    months = ListProperty(["Januar", "Februar", "Mars", "April", "Mai", "Juni", "Juli", "August", "September", "Oktober", "November", "Desember"])

    mailContainers = ListProperty(None)

    def playSound(self):
        winsound.PlaySound("God_morgen.wav", winsound.SND_FILENAME | winsound.SND_ASYNC)

    def update(self, dt):
        self.time = datetime.datetime.now()

        if (self.time - self.waited).seconds >= 60:
            self.page = 0

        weekday = self.time.weekday()
        if weekday != self.currentDay:
            self.currentDay = weekday
            #self.updateSchedule()
            #ukeplan = convert_from_path("ukeplan.pdf")
            self.updateImages()

        if self.time.hour != self.currentHour:
            self.currentHour = self.time.hour
            self.updateWeather()
            self.update_mails()

    def updateSchedule(self):
        if self.currentDay < 5:
            current = self.schedulePlans[self.currentDay]
        else:
            current = self.schedulePlans[0]

        self.schedulePlan.clear_widgets()
        self.schedulePlan.add_widget(Label(text="Timeplan", font_size=60))
        for i in current:
            self.schedulePlan.add_widget(Label(text=i, font_size=30))

    def updateWeather(self):
        weather = Yr(location_name="Norge/Oslo/Oslo/Kringsjå")
        now = weather.now()
        self.weatherSymbol = now["symbol"]["@number"]
        self.degrees = int(now["temperature"]["@value"])

    def updateImages(self):

        mailbox = account.mailbox()
        query = mailbox.new_query().on_attribute('subject').contains('Melding fra portalen: Informasjon uke').order_by("subject", ascending=False)
        message = mailbox.get_message(query=query, download_attachments=True)
        message.attachments[0].save("./", "ukeplan.pdf")

        inbox = mailbox.inbox_folder()
        messages = inbox.get_messages()

        pdf2jpeg("./ukeplan.pdf", "./ukeplan")

        page1 = pygame.image.load("ukeplan1.png")
        page2 = pygame.image.load("ukeplan2.png")

        w = 965
        h = 555
        x = 141
        y = 984

        timeplan = pygame.surface.Surface((w, h))
        timeplan.blit(page1, (0, 0), (x, y, w, h))

        pygame.image.save(timeplan, "timeplan.png")

        w = 907
        h = 1400
        x = 141
        y = 220

        lekser = pygame.surface.Surface((w, h))
        lekser.blit(page2, (0, 0), (x, y, w, h))

        w2 = lekser.get_height() - 1
        while lekser.get_at((500, w2))[0] > 100 and w2 > 0:
            w2 -= 1

        y2 = 0
        while lekser.get_at((500, y2))[0] > 100 and y2 < lekser.get_height() - 1:
            y2 += 1

        lekser2 = pygame.surface.Surface((w, w2 - y2))
        lekser2.blit(lekser, (0, 0), (0, y2, w, w2))

        pygame.image.save(lekser2, "lekser.png")

        self.timeplanImage.reload()
        self.lekserImage.reload()

    def on_touch_down(self, touch):
        self.waited = datetime.datetime.now()

        if self.page == 0 and touch.is_double_tap:
            self.alarm = not self.alarm
            if self.alarm:
                self.alarmText.color = (0, 1, 0, 1)
            else:
                self.alarmText.color = (1, 0, 0, 1)

        return super().on_touch_down(touch)

    def update_mails(self):
        mailbox = account.mailbox()
        messages = mailbox.get_messages()

        

        #[x.extract() for x in soup.findAll('a')]
        #print(soup.get_text())
        #print(soup.prettify())

        self.mails.clear_widgets()
        j = 0
        for i in messages:
            if i.sender.address == "noreply@portal.skoleplattform.no" and j < 5:
                self.mails.add_widget(Label(text=i.subject[22:]))
                soup = BeautifulSoup(i.body, "html.parser")
                divs = soup.findAll("div")
                divs[len(divs) - 1].extract()

                text = soup.get_text()
                print(text)
                text = addLines(text, 100)
                #print(text)

                self.mailContainers[j].add_widget(Label(text=text))

                j += 1

def pdf2jpeg(pdf_input_path, jpeg_output_path):
    args = ["pdf2jpeg", # actual value doesn't matter
            "-dNOPAUSE",
            "-sDEVICE=jpeg",
            "-r144",
            "-sOutputFile=" + jpeg_output_path + "%d.png",
            pdf_input_path]
    encoding = locale.getpreferredencoding()
    ghostscript.Ghostscript(*[a.encode(encoding) for a in args])

def addLines(string, max_characters):
    string += " trash"
    lines = []
    last_space = 0
    total_index = 0
    line_index = 0
    while total_index + line_index < len(string):
        #print(line_index, total_index, last_space, string[total_index + line_index], lines, len(string))
        if string[total_index + line_index] ==  " ":
            last_space = total_index + line_index
        if string[total_index + line_index:total_index + line_index + 1] == "\n":
            lines.append(string[total_index:total_index + line_index])
            total_index += line_index + 1
            line_index = 0
            last_space = 0
        if line_index >= max_characters or total_index + line_index == len(string):
            lines.append(string[total_index:last_space])
            line_index = 0
            total_index = last_space + 1
            last_space = 0
        else:
            line_index += 1
        
    return '\n'.join(lines)

#print(addLines("""SMS-melding
"""
Av: Sigurd Enge, Dato: 04.12.2018 kl. 11:27

Til: 8SA-8D Samfunnsfag (Elever + Lærere)
 

Nordberg skole:
Hei, 8D! Nå har jeg lagt ut en plan for vurderingen i samf på its. Sjekk kriteriene. Vi starter på torsdag. Mvh, Sig
SMSen kan ikke besvares"""#, 10))

class PagesApp(App):
    def build(self):
        #Config.set("graphics", "fullscreen", "auto")
        pages = PagesClass()
        Clock.schedule_interval(pages.update, 1)
        return pages

PagesApp().run() 