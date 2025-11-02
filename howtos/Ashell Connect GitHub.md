Title: a-shell connect GitHub

Tags: A-shell, GitHub, HowTo, Instructions 

Intro: I have had a hard time in the past trying to make .txt files and being able to work on simple files on GitHub repos from my phone and ipad. There are times that I wanted to just write a quick journal entry or work on a file on my repo while I was waiting for a friend to show for a lunch hangout or when I was just bored waiting. I am excited that this exist. 

a-shell is an app in the app store that I found recently due to someone saying that they only had their ipad for a workshop on GitHub. In the instructions of the workshop, we ask them to use a computer, but this happens more often than not. So I went on the search for a Github ipad GUI. At first I found one that looked amazing it had all the terms and did almost all the same items as GitHub Desktop but they charged you to push... I was so disappointed, then I found another one that let you do 3 pushes for free then you had to pay. I was slight less disappointed. Yet this got me thinking and searching to see if can do it with command line for free. Which is when I found a-shell this is where the headache started for a different reason. I needed to set up the connection to pull and push to GitHub. I was unaware of what my next 8 hours would look like.

Side note scene set up time (This is for nothing about a-shell, I just wanted to write this): It was a Friday night, my cousin is having a small birthday party for her now 7 year old child. She decided to hire a magician to do some magic which actually entertained not just the small children but all of the adults there, there I was sitting on my phone half looking up at each magic trick trying to connect a-shell to GitHub and as I am silently cussing as myself, "I did make it, why is it not working." I keep getting the same error, the .ssh does not exist, when I did the cat copy code to put it into GitHubs settings. 

So once I was at the house, I was able to really research what was going on which is when I found the forum post (I did try chatgpt by the way, chatgpt tells you just go to the .ssh folder or here is the code, it will work. Where as a forum post will say, dude I know I went through the same struggles, here is everything I have tried and this is how I fixed it. This is why I will always appreciate forums more than a chatbot) that told me what I was missing, which is exactly where the .ssh folder is on the ipad, when the pwd is semi lying to you... (it was not, but is was not picking it up with the CAT code and when you went looking for it is was like nope I am not here, until after 8 hours where its like "AHA YOU FOUND ME!" because you had just decided to write the same cd then press tab not in the root folder but in the folder you where in the entire time. So below are the step for future me so I do not forget

Steps:
1. Open a-shell 
2. install git
> pkg install git
3. check with ls that there is a bin folder
> ls
4. also check what commands you can do with lg2
> lg2
5. Clone git repo from GitHub, you will need url of the repo you are cloning, there is no big green button on phone app so go to repo click the circle with the three dots and click share then just click the option copy, you will have to add ".git" at the end.
> git clone "GitHub Repo URL with .git at the end"
6. check with ls that your repo is there
> ls
7. Now generate the public/private key
> ssh-keygen -t rsa -b 4096 -C "your email you use on GitHub"
8. click enter for default file name and file folder, unless one already exist click Y first, then for the passwords click enter then enter (You can make your first protected with passcodes, I did...) It will give you the path its saved at for both private and public and display the fingerprint and text and graphic
9. In the folder where you have the bin and the git clone, write cd then hit tab then type a dot then "s" the .ssh/ should come up if it doesn't, you will probably have to have an extra set of fun searching the answer for next 8 hours, I reccomend going to a magic show first. But it should of came up, I have tested it many times. Go into that folder
> cd .ssh/
10. check that the keys are in there with ls
> ls
11. display the public key and manually copy it
> head id_ed25519.pub
12. go to GitHub on a browser (I used Firefox on my ipad and phone), navigate to the setting then to SSH and GPG keys, click the green button that says "New SSH key" that call it the device and then paste the manually copied public key. This is a good point to make a GPG key you will need it when you git push to the repo.
13. Make the GPG key, IMPORTANT: IT WILL ONLY DISPLAY THE TEXT IT GIVES YOU FOR THE GPG KEY ONCE, MAKE SURE TO COPY AND SAVE IT SOMEWHERE WHERE IT IS SAFE AND ACCESSABLE. I have a friend who told me it was a pain replacing the old one, you do not want to go through that, right?
14. Okay finally! Let's test it! Navigate to the git repo folder and make sure you are inside of it.
> cd "repo folder"
15. make a test file, we will push it to the repo
> vim testfile.txt
16. pull the git repo
> git pull
17. add the file to be uploaded to git
> git add .
18. look at the status, it should say changes to commit or something similar
> git status
19. make a commit message (whisper: you will not be able to do this, it will say you will have an extra step, the first time you do this, thank goodness for testing)
> git commit -m "Your message here"
20. Oh look an error, I would never have guessed. Add your user.name and User.Email from GitHub.
> lg2 config user.name 'Your Name'
21. Add your user.email from GitHub now
> lg2 config user.email youremail@example.com
22. Redo the commit message
> git commit -m "Your message here"
23. check the status, I only get back what branch it's on.
> git status
24. Now push to the repo
> git push
25. It will ask for the password, this is going to be the GPG key not your GitHub password, I just copy and pasted in there. (I tried my GitHub password for a good 10 attempts before I remembered to use the GPG key (if it gives you a error that says "Error Pushing [-1] - unexpected EOF", they go buy a ticket to the local magic show... it will be another 8 hours)
26. Success *chit's voice* (tiktok reference), if it says pushed. If for some odd reason you get the error "Error Pushing [-1] - unexpected EOF" or something else as an error... do you like card tricks? I know a magician you might like. Over played joke... yeah I know, but you now have a inside joke with me (or yourself). But go check that it is on your GitHub Repo Online.

Summary: Well that is how you connect a-shell to your GitHub. To future me, I know you are going to want to thank past you (me) for taking the time to write this out, buy me a coffee or lunch if you want.

Time is takes to set-up if successful: 5-10 minutes 

Links to the forums that I found useful:
"https://forum.obsidian.md/t/mobile-automatic-sync-with-github-on-ios-for-free-via-a-shell/46150?page=2"
it has links to other websites including the main GitHub Documentation on SSH, it is for Obsindian app, but it gave me the information to fix my issue.

