from twitchio.ext import commands
import pandas as pd
from datetime import datetime as dt, timedelta as td
from threading import Thread
from keyboard import Keyboard
import time

class Bot(commands.Bot):

    def __init__(self):
        self.time_called = 0
        self.start_time = 0
        self.end_time = 0
        self.is_blackout_running = False
        self.blackout_Thread = Thread(target = self.run_timer)
        self.blackout_seconds = 0
        self.df_users = pd.read_excel('Blackout.xlsx',sheet_name = "Blackout")
        self.df_timer =  pd.read_excel('Timer.xlsx',sheet_name = "Timer")
        # Initialise our Bot with our access token, prefix and a list of channels to join on boot...
        # prefix can be a callable, which returns a list of strings or a string...
        # initial_channels can also be a callable which returns a list of strings...
        super().__init__(nick="Pungoro",token= , prefix='!', initial_channels=[''])
        

    async def event_ready(self):
        # Notify us when everything is ready!
        # We are logged in and ready to chat and use commands...
        self.blackout_Thread.start()
        print("User's printing")
        print(f'Logged in as | {self.nick}')
        print(f'User id is | {self.user_id}')

    
    @commands.command()
    async def blackout(self,ctx:commands.Context):
        if not str(ctx.author.name) in self.df_users.UserName.values:
            self.time_called = pd.Timestamp.now()
            time_start = self.time_called.time()
            self.end_time = self.time_called + td(seconds=10)
            self.df_users.loc[len(self.df_users.index)] = [str(ctx.author.name),self.time_called,self.end_time,True,True]
            self.df_users.to_excel('Blackout.xlsx',sheet_name = "Blackout", index=False)
            self.blackout_seconds += 10
            await ctx.send("Hello {}, you have extended the blackout screen to {} seconds".format(ctx.author.name,self.blackout_seconds))
            
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
                await ctx.send("Hello {}, you have extended the blackout screen to {} seconds".format(ctx.author.name,self.blackout_seconds))
                
        else:
            self.save_log(ctx.author.name)
            await ctx.send("Hello {}, you have extended the blackout screen to {} seconds".format(ctx.author.name,self.blackout_seconds))
        
        if not self.is_blackout_running:
            
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
            # self.blackout_seconds += 10
            # await ctx.send("Hello {}, you have extended the blackout screen to {} seconds".format(ctx.author.name,self.blackout_seconds))

        

    def save_log(self,username):
            self.blackout_seconds += 10
            self.time_called = pd.Timestamp.now()
            time_start = self.time_called.time()
            self.end_time= self.time_called + td(seconds=10)
            # print("Computed end time: {}".format(self.end_time))
            self.df_users.loc[self.df_users.UserName == username] = [username,self.time_called,self.end_time,True,True]
            self.df_users.to_excel('Blackout.xlsx',sheet_name = "Blackout", index=False)
            

    
    def run_timer(self):
        while 1:
            while self.is_blackout_running:
                beg_time = self.start_time.time()
                end_time = self.end_time.time()

                self.ScreenOff()
                print("Screen is Off")
                while beg_time < end_time:
                    beg_time = dt.now().time()
                    end_time = self.end_time.time()
                # print("Finished at {}".format(dt.now()))
                self.ScreenOn()
                self.blackout_seconds= 0
                # print("Screen is On")
                self.is_blackout_running = False
                print("end time was: {}".format(end_time))
                print("Start time was {}".format(self.start_time.time()))

                print("loop finished")
        
            while not self.is_blackout_running:
                # print("It's still running")
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