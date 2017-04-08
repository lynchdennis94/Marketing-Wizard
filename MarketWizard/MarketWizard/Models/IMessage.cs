using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MarketWizard.Models {
    interface IMessage {
        public List<String> getRecepients();
        public String getBody();
    }
}
