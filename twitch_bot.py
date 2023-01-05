from twitchio.ext import commands
import pandas as pd
from datetime import datetime as dt, timedelta as td
from threading import Thread
from keyboard import Keyboard
from sound import Sound
import time

class Bot(commands.Bot):

    def __init__(self):
        self.time_called = 0
        self.start_time = 0
        self.end_time = 0
        self.is_blackout_running = False
        self.is_soundproof_running = False
        self.can_blackout_run = True
        self.can_soundproof_run = True
        self.blackout_Thread = Thread(target = self.run_timer)
        self.effects_seconds = 0
        self.df_users = pd.read_excel('Blackout.xlsx',sheet_name = "Blackout")
        self.df_timer =  pd.read_excel('Timer.xlsx',sheet_name = "Timer")
        # Initialise our Bot with our access token, prefix and a list of channels to join on boot...
        # prefix can be a callable, which returns a list of strings or a string...
        # initial_channels can also be a callable which returns a list of strings...
        super().__init__(nick="Pungoro",token=, prefix='!', initial_channels=['Pungoro'])
        

    async def event_ready(self):
        # Notify us when everything is ready!
        # We are logged in and ready to chat and use commands...
        self.blackout_Thread.start()
        print("User's printing")
        print(f'Logged in as | {self.nick}')
        print(f'User id is | {self.user_id}')

    
    @commands.command()
    async def blackout(self,ctx:commands.Context):
        if self.can_blackout_run:
            if not str(ctx.author.name) in self.df_users.UserName.values:
                self.time_called = pd.Timestamp.now()
                time_start = self.time_called.time()
                self.end_time = self.time_called + td(seconds=10)
                self.df_users.loc[len(self.df_users.index)] = [str(ctx.author.name),self.time_called,self.end_time,True,True]
                self.df_users.to_excel('Blackout.xlsx',sheet_name = "Blackout", index=False)
                self.effects_seconds += 10
                await ctx.send("Hello {}, you have extended the blackout screen to {} seconds".format(ctx.author.name,self.effects_seconds))

            elif self.df_users.loc[self.df_users.UserName == ctx.author.name].Cooldown.values[0]:
                self.end_time = self.df_users.loc[self.df_users.UserName == ctx.author.name].Time_End.values[0]
                self.end_time = pd.to_datetime(self.end_time,errors='ignore')
                wait_end = self.end_time + td(seconds = 10)
                if not wait_end < dt.now():
                    time_remaining = wait_end - dt.now() 
                    await ctx.send("Hello {}, please wait {} seconds to send the command again".format(ctx.author.name,time_remaining.seconds))
                    return
                else:
                    self.save_log(ctx.author.name)
                    await ctx.send("Hello {}, you have extended the blackout screen to {} seconds".format(ctx.author.name,self.effects_seconds))

            else:
                self.save_log(ctx.author.name)
                await ctx.send("Hello {}, you have extended the blackout screen to {} seconds".format(ctx.author.name,self.effects_seconds))

            if not self.is_blackout_running:
                self.can_soundproof_run = False
                self.start_time = self.time_called
                self.df_timer.loc[len(self.df_timer.index)] = [self.start_time,self.end_time]
                self.df_timer.to_excel('Timer.xlsx',sheet_name = "Timer", index=False)
                # print("Start loop through time in different thread")
                # print("Original end time = {}".format(self.end_time.time()))
                self.is_blackout_running = True
                # if not self.blackout_Thread.is_alive():
                #     print("It's not alive")


            else:
                self.end_time = self.end_time+ td(seconds=10)
                self.df_timer.Start_Time[0] = self.start_time
                self.df_timer.End_Time[0]= self.end_time
                self.df_timer.to_excel('Timer.xlsx',sheet_name = "Timer", index=False)
                # print("Just update the end time")
                # print("Updated end time = {}".format(self.end_time.time()))
                # self.effects_seconds += 10
                # await ctx.send("Hello {}, you have extended the blackout screen to {} seconds".format(ctx.author.name,self.effects_seconds))
        else:
            await ctx.send("Hello {}, please wait for the soundproof session to finish".format(ctx.author.name,self.effects_seconds))
        


    @commands.command()
    async def soundproof(self,ctx:commands.Context):
        if self.can_soundproof_run:
            if not str(ctx.author.name) in self.df_users.UserName.values:
                self.time_called = pd.Timestamp.now()
                time_start = self.time_called.time()
                self.end_time = self.time_called + td(seconds=10)
                self.df_users.loc[len(self.df_users.index)] = [str(ctx.author.name),self.time_called,self.end_time,True,True]
                self.df_users.to_excel('Blackout.xlsx',sheet_name = "Blackout", index=False)
                self.effects_seconds += 10
                await ctx.send("Hello {}, you have extended the soundproof session to {} seconds".format(ctx.author.name,self.effects_seconds))

            elif self.df_users.loc[self.df_users.UserName == ctx.author.name].Cooldown.values[0]:
                self.end_time = self.df_users.loc[self.df_users.UserName == ctx.author.name].Time_End.values[0]
                self.end_time = pd.to_datetime(self.end_time,errors='ignore')
                wait_end = self.end_time + td(seconds = 10)
                if not wait_end < dt.now():
                    time_remaining = wait_end - dt.now() 
                    await ctx.send("Hello {}, please wait {} seconds to send the command again".format(ctx.author.name,time_remaining.seconds))
                    return
                else:
                    self.save_log(ctx.author.name)
                    await ctx.send("Hello {}, you have extended the soundproof session to {} seconds".format(ctx.author.name,self.effects_seconds))

            else:
                self.save_log(ctx.author.name)
                await ctx.send("Hello {}, you have extended the soundproof session to {} seconds".format(ctx.author.name,self.effects_seconds))

            if not self.is_soundproof_running:
                self.can_blackout_run = False
                self.start_time = self.time_called
                self.df_timer.loc[len(self.df_timer.index)] = [self.start_time,self.end_time]
                self.df_timer.to_excel('Timer.xlsx',sheet_name = "Timer", index=False)
                # print("Start loop through time in different thread")
                # print("Original end time = {}".format(self.end_time.time()))
                self.is_soundproof_running = True
                # if not self.blackout_Thread.is_alive():
                #     print("It's not alive")


            else:
                self.end_time = self.end_time+ td(seconds=10)
                self.df_timer.Start_Time[0] = self.start_time
                self.df_timer.End_Time[0]= self.end_time
                self.df_timer.to_excel('Timer.xlsx',sheet_name = "Timer", index=False)
                # print("Just update the end time")
                # print("Updated end time = {}".format(self.end_time.time()))
                # self.effects_seconds += 10
                # await ctx.send("Hello {}, you have extended the blackout screen to {} seconds".format(ctx.author.name,self.effects_seconds))
        else:
            await ctx.send("Hello {}, please wait for the blackout session to finish".format(ctx.author.name,self.effects_seconds))

    def save_log(self,username):
            self.effects_seconds += 10
            self.time_called = pd.Timestamp.now()
            time_start = self.time_called.time()
            self.end_time= self.time_called + td(seconds=10)
            # print("Computed end time: {}".format(self.end_time))
            self.df_users.loc[self.df_users.UserName == username] = [username,self.time_called,self.end_time,True,True]
            self.df_users.to_excel('Blackout.xlsx',sheet_name = "Blackout", index=False)
            

    
    def run_timer(self):
        while 1:
            if not self.can_soundproof_run:
                while self.is_blackout_running:
                    beg_time = self.start_time.time()
                    end_time = self.end_time.time()

                    self.ScreenOff()
                    while beg_time < end_time:
                        beg_time = dt.now().time()
                        end_time = self.end_time.time()
                    # print("Finished at {}".format(dt.now()))
                    self.ScreenOn()
                    self.effects_seconds= 0
                    self.is_blackout_running = False
                    print("end time was: {}".format(end_time))
                    print("Start time was {}".format(self.start_time.time()))

                    self.can_soundproof_run = True
            elif not self.can_blackout_run:
                while self.is_soundproof_running:
                    beg_time = self.start_time.time()
                    end_time = self.end_time.time()

                    Sound.volume_set(0)
                    while beg_time < end_time:
                        beg_time = dt.now().time()
                        end_time = self.end_time.time()
                    Sound.volume_set(70)
                    self.effects_seconds= 0
                    self.is_soundproof_running = False
                    print("end time was: {}".format(end_time))
                    print("Start time was {}".format(self.start_time.time()))

                    self.can_blackout_run = True


            while not self.is_blackout_running and not self.is_soundproof_running:
                pass
                

    def ScreenOff(self):
        Keyboard.keyDown(Keyboard.VK_LWIN)
        Keyboard.keyDown(Keyboard.VK_LSHIFT)
        Keyboard.keyDown(Keyboard.VK_LEFT)
        Keyboard.keyUp(Keyboard.VK_LWIN)
        Keyboard.keyUp(Keyboard.VK_LSHIFT)
        Keyboard.keyUp(Keyboard.VK_LEFT)

        # print("Turning Screen Off")
        # ctypes.windll.user32.SendMessageW(65535, 274, 61808, 2)
        return

    def ScreenOn(self):
        Keyboard.keyDown(Keyboard.VK_LWIN)
        Keyboard.keyDown(Keyboard.VK_LSHIFT)
        Keyboard.keyDown(Keyboard.VK_RIGHT)
        Keyboard.keyUp(Keyboard.VK_LWIN)
        Keyboard.keyUp(Keyboard.VK_LSHIFT)
        Keyboard.keyUp(Keyboard.VK_RIGHT)

        # print("Turning Screen On")
        # ctypes.windll.user32.SendMessageW(65535, 274, 61808, -1)
        # win32api.mouse_event(win32con.MOUSEEVENTF_MOVE, 0, 0)
        return
    
bot = Bot()
bot.run()
