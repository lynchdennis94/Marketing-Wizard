using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketWizard.Utilities
{
    // This tracker will display small messages about the status of the project
    class StatusTracker
    {
        private int totalObjective; // The maximum number of things to be done
        private int currentObjective; // Where the user is currently
        private String currentStatus;

        public StatusTracker()
        {
            currentObjective = 0;
            totalObjective = 0;
            currentStatus = "";
        }

        public void setObjectiveTotal(int total)
        {
            totalObjective = total;
        }

        public void incrementCurrentObjective()
        {
            currentObjective++;
        }

        public void setStatus(String newStatus)
        {
            currentStatus = newStatus;
        }

        public String getObjectiveStatus()
        {
            return currentStatus + " (" + (currentObjective*100/totalObjective) + "% )";
        }
    }
}
