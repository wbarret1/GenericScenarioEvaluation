namespace GenericScenarioEvaluation
{
    using System;
    using System.Data.Entity;
    using System.Linq;

    public class GenericScenarioModel : DbContext
    {
        // Your context has been configured to use a 'GenericScenarioModel' connection string from your application's 
        // configuration file (App.config or Web.config). By default, this connection string targets the 
        // 'GenericScenarioEvaluation.GenericScenarioModel' database on your LocalDb instance. 
        // 
        // If you wish to target a different database and/or database provider, modify the 'GenericScenarioModel' 
        // connection string in the application configuration file.
        public GenericScenarioModel()
            : base("name=GenericScenarioModel")
        {
        }

        // Add a DbSet for each entity type that you want to include in your model. For more information 
        // on configuring and using a Code First model, see http://go.microsoft.com/fwlink/?LinkId=390109.

        public virtual DbSet<GenericScenario> GenericsScenarios { get; set; }
        public virtual DbSet<OccupationalExposure> OccupationalExposures { get; set; }
        public virtual DbSet<EnvironmentalRelease> EnvironmentalReleases { get; set; }
        public virtual DbSet<ControlTechnology> ControlTechnologies { get; set; }
        public virtual DbSet<ProcessDescription> ProcessDescriptions { get; set; }
        public virtual DbSet<ProductionRate> ProductionRates { get; set; }
        public virtual DbSet<RemainingValue> RemainingValues { get; set; }
        public virtual DbSet<DataElement> DataElements { get; set; }
        public virtual DbSet<Source> Sources { get; set; }
        public virtual DbSet<Shift> Shifts { get; set; }
        public virtual DbSet<Site> Sites { get; set; }
        public virtual DbSet<UseRate> UseRates { get; set; }
        public virtual DbSet<Worker> Workers { get; set; }
    }

    //public class MyEntity
    //{
    //    public int Id { get; set; }
    //    public string Name { get; set; }
    //}
}