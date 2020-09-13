namespace SpeechToTextWPFSample.Models
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class MyDatabase : DbContext
    {
        public MyDatabase()
            : base("name=MyDatabase")
        {
        }

        public virtual DbSet<Category> Categories { get; set; }
        public virtual DbSet<CustomIntelliSense> CustomIntelliSenses { get; set; }
        public virtual DbSet<Language> Languages { get; set; }
        public virtual DbSet<Launcher> Launchers { get; set; }
        public virtual DbSet<MultipleLauncher> MultipleLaunchers { get; set; }
        public virtual DbSet<Values_to_Insert> Values_to_Inserts { get; set; }
        public virtual DbSet<ApplicationsToKill> ApplicationsToKill { get; set; }
        public virtual DbSet<Computer> Computers { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {



        }
    }
}
