using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpeechToTextWPFSample.Models
{
    [Table("ApplicationsToKill")]
    public class ApplicationsToKill
    {
        public int Id { get; set; }

        [Required]
        public string ProcessName { get; set; }

        [Required]
        public string CommandName { get; set; }

        public bool Display { get; set; }
    }
}
