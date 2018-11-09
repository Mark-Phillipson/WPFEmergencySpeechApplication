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

        public virtual DbSet<tblCategory> tblCategories { get; set; }
        public virtual DbSet<CustomIntelliSense> tblCustomIntelliSenses { get; set; }
        public virtual DbSet<tblLanguage> tblLanguages { get; set; }
        public virtual DbSet<tblLauncher> tblLaunchers { get; set; }
        public virtual DbSet<tblMultipleLauncher> tblMultipleLaunchers { get; set; }
        public virtual DbSet<Values_to_Insert> Values_to_Inserts { get; set; }
        public virtual DbSet<ApplicationsToKill> ApplicationsToKill { get; set; }
        public virtual DbSet<Computer> Computers { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {

            modelBuilder.Entity<tblCategory>()
                .HasMany(e => e.tblCustomIntelliSenses)
                .WithOptional(e => e.tblCategory)
                .HasForeignKey(e => e.Category_ID);

            modelBuilder.Entity<tblCategory>()
                .HasMany(e => e.tblLaunchers)
                .WithOptional(e => e.tblCategory)
                .HasForeignKey(e => e.Menu);

            modelBuilder.Entity<CustomIntelliSense>()
                .Property(e => e.SSMA_TimeStamp)
                .IsFixedLength();

            modelBuilder.Entity<tblLanguage>()
                .HasMany(e => e.tblCustomIntelliSenses)
                .WithOptional(e => e.tblLanguage)
                .HasForeignKey(e => e.Language_ID);

            modelBuilder.Entity<Values_to_Insert>()
                .Property(e => e.RowVersion)
                .IsFixedLength();


        }
    }
}
