﻿//------------------------------------------------------------------------------
// <auto-generated>
//     Este código se generó a partir de una plantilla.
//
//     Los cambios manuales en este archivo pueden causar un comportamiento inesperado de la aplicación.
//     Los cambios manuales en este archivo se sobrescribirán si se regenera el código.
// </auto-generated>
//------------------------------------------------------------------------------

using System;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;

public partial class CertelEntities : DbContext
{
    public CertelEntities()
        : base("name=CertelEntities")
    {
    }

    protected override void OnModelCreating(DbModelBuilder modelBuilder)
    {
        throw new UnintentionalCodeFirstException();
    }

    public virtual DbSet<Aparato> Aparato { get; set; }
    public virtual DbSet<Caracteristica> Caracteristica { get; set; }
    public virtual DbSet<Cliente> Cliente { get; set; }
    public virtual DbSet<Cumplimiento> Cumplimiento { get; set; }
    public virtual DbSet<EquipoUtilizado> EquipoUtilizado { get; set; }
    public virtual DbSet<Especificos> Especificos { get; set; }
    public virtual DbSet<EstadoInspeccion> EstadoInspeccion { get; set; }
    public virtual DbSet<EstadoServicio> EstadoServicio { get; set; }
    public virtual DbSet<EstructuraInfome> EstructuraInfome { get; set; }
    public virtual DbSet<Evaluacion> Evaluacion { get; set; }
    public virtual DbSet<FormatoNorma> FormatoNorma { get; set; }
    public virtual DbSet<MenuAcceso> MenuAcceso { get; set; }
    public virtual DbSet<MenuModulo> MenuModulo { get; set; }
    public virtual DbSet<Norma> Norma { get; set; }
    public virtual DbSet<Pda> Pda { get; set; }
    public virtual DbSet<Requisito> Requisito { get; set; }
    public virtual DbSet<Rol> Rol { get; set; }
    public virtual DbSet<Servicio> Servicio { get; set; }
    public virtual DbSet<TerminosYDefiniciones> TerminosYDefiniciones { get; set; }
    public virtual DbSet<TipoFuncionamientoAparato> TipoFuncionamientoAparato { get; set; }
    public virtual DbSet<TipoNorma> TipoNorma { get; set; }
    public virtual DbSet<Titulo> Titulo { get; set; }
    public virtual DbSet<Usuario> Usuario { get; set; }
    public virtual DbSet<UsuarioRol> UsuarioRol { get; set; }
    public virtual DbSet<ValoresEspecificos> ValoresEspecificos { get; set; }
    public virtual DbSet<InspeccionNorma> InspeccionNorma { get; set; }
    public virtual DbSet<Fotografias> Fotografias { get; set; }
    public virtual DbSet<ActualizacionesPendientes> ActualizacionesPendientes { get; set; }
    public virtual DbSet<CustomInforme> CustomInforme { get; set; }
    public virtual DbSet<Informe> Informe { get; set; }
    public virtual DbSet<Inspeccion> Inspeccion { get; set; }
    public virtual DbSet<ObservacionTecnica> ObservacionTecnica { get; set; }
    public virtual DbSet<TipoInforme> TipoInforme { get; set; }
    public virtual DbSet<DestinoProyecto> DestinoProyecto { get; set; }
    public virtual DbSet<FotografiaTecnica> FotografiaTecnica { get; set; }
    public virtual DbSet<NormasAsociadas> NormasAsociadas { get; set; }
}
