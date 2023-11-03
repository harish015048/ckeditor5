using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.SessionState;

namespace CMCai.Actions
{
    public class SessionDb
    {
        public static Dictionary<string, HttpSessionState> sessionDB = new Dictionary<string, HttpSessionState>();
        public SessionDb()
        {
        }

        public static void addUserAndSession(string name, HttpSessionState session)
        {
            sessionDB.Add(name, session);
        }

        public static bool containUser(string name)
        {
            //in fact,here I also want to check if this session is active or not ,but I do not find the method like session.isActive() or session.HasExpire() or something else.

            return sessionDB.ContainsKey(name);
        }

        public static void removeUser(string name)
        {
            if (sessionDB.ContainsKey(name))
            {
                sessionDB.Remove(name);
            }
        }
    }
}