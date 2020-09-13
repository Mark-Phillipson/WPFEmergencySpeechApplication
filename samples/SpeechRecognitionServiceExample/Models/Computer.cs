using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpeechToTextWPFSample.Models
{
    public class Computer
    {
        public Computer()
        {
            CustomIntelliSenses = new HashSet<CustomIntelliSense>();
            Launchers = new HashSet<Launcher>();
        }
        public int ID { get; set; }

        [Required]
        [Display(Name ="Computer Name")]
        public string ComputerName { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<CustomIntelliSense> CustomIntelliSenses { get; set; }

        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<Launcher> Launchers { get; set; }
    }
}
