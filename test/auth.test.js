'use strict';
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
function load(pca) {
 const context = {self:{}, msal:{createNestablePublicClientApplication:async()=>pca},setTimeout,clearTimeout,Office:{context:{mailbox:{userProfile:{emailAddress:'sample@example.org'}},roamingSettings:{get(){return false},set(){},saveAsync(fn){fn()}}}}};
 vm.runInNewContext(fs.readFileSync(require('node:path').join(__dirname,'../src/graph.js'),'utf8'),context);
 return context.self.GraphData;
}
(async()=>{
 // The certification tenant denied /me/memberOf with group scopes alone.
 // Exercise both token paths: a cached token must be upgraded too.
 for (const interactive of [false, true]) {
   const requests=[];
   const auth=load({getAllAccounts:()=>[],
     acquireTokenSilent:async request=>{requests.push(request);if(interactive)throw Error('consent required');return {accessToken:'membership-token'}},
     acquireTokenPopup:async request=>{requests.push(request);return {accessToken:'membership-token'}}});
   assert.equal(await auth.getToken(),'membership-token');
   assert.equal(requests.length,interactive?2:1);
   for(const request of requests) assert.ok(request.scopes.includes('User.Read'),'membership calls need the signed-in user scope');
 }
 let popups=0;
 const g=load({getAllAccounts:()=>[],acquireTokenSilent:async()=>({accessToken:'token',account:{username:'sample@example.org'}}),acquireTokenPopup:async()=>{popups++;return {accessToken:'interactive'}}});
 assert.equal(await g.getToken(),'token');assert.equal(await g.currentAccount(),'sample@example.org');
 await g.signOut();assert.equal(await g.currentAccount(),null);assert.equal(await g.getToken(),'interactive');assert.equal(popups,1);
 let finish;let entered; const started=new Promise(r=>entered=r);
 const race=load({getAllAccounts:()=>[],acquireTokenSilent:()=>{entered();return new Promise(r=>finish=r)},acquireTokenPopup:async()=>{throw Error('must not open popup after sign out')}});
 const pending=race.getToken();await started;await race.signOut();finish({accessToken:'stale'});await assert.rejects(pending,/cancelled/);assert.equal(await race.currentAccount(),null);
 let completePopup; let popupEntered;
 const popupStarted = new Promise(r=>popupEntered=r);
 const popupRace=load({getAllAccounts:()=>[],acquireTokenSilent:async()=>{throw Error('interaction required')},acquireTokenPopup:()=>{popupEntered();return new Promise(r=>completePopup=r)}});
 const popupPending=popupRace.getToken(); await popupStarted; await popupRace.signOut(); completePopup({accessToken:'stale popup'});
 await assert.rejects(popupPending,/cancelled/); assert.equal(popupRace.isSignedOut(),true);
 console.log('Authentication state and sign-out race regressions passed.');
})().catch(e=>{console.error(e);process.exitCode=1});
