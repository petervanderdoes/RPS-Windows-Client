Imports System
Imports System.ComponentModel.DataAnnotations.Schema
Imports System.Data.Common
Imports System.Data.Entity
Imports System.Data.SQLite
Imports System.Linq
Imports System.Runtime.InteropServices

Namespace Entities


    <ComVisible(False)>
    Partial Public Class rpsEntities
        Inherits DbContext

        Public Sub New(c As String)
            MyBase.New(New SQLiteConnection() With {.ConnectionString = c}, True)
        End Sub

        Public Overridable Property awards As DbSet(Of award)
        Public Overridable Property classifications As DbSet(Of classification)
        Public Overridable Property clubs As DbSet(Of club)
        Public Overridable Property club_award As DbSet(Of club_award)
        Public Overridable Property club_classification As DbSet(Of club_classification)
        Public Overridable Property club_medium As DbSet(Of club_medium)
        Public Overridable Property CompetitionEntries As DbSet(Of CompetitionEntry)
        Public Overridable Property media As DbSet(Of medium)

        Protected Overrides Sub OnModelCreating(ByVal modelBuilder As DbModelBuilder)
            modelBuilder.Entity(Of award)() _
                .HasMany(Function(e) e.club_award) _
                .WithOptional(Function(e) e.award) _
                .HasForeignKey(Function(e) e.award_id)

            modelBuilder.Entity(Of classification)() _
                .HasMany(Function(e) e.club_classification) _
                .WithOptional(Function(e) e.classification) _
                .HasForeignKey(Function(e) e.classification_id)

            modelBuilder.Entity(Of club)() _
                .HasMany(Function(e) e.club_award) _
                .WithOptional(Function(e) e.club) _
                .HasForeignKey(Function(e) e.club_id)

            modelBuilder.Entity(Of club)() _
                .HasMany(Function(e) e.club_classification) _
                .WithOptional(Function(e) e.club) _
                .HasForeignKey(Function(e) e.club_id)

            modelBuilder.Entity(Of club)() _
                .HasMany(Function(e) e.club_medium) _
                .WithOptional(Function(e) e.club) _
                .HasForeignKey(Function(e) e.club_id)

            modelBuilder.Entity(Of medium)() _
                .HasMany(Function(e) e.club_medium) _
                .WithOptional(Function(e) e.medium) _
                .HasForeignKey(Function(e) e.medium_id)
        End Sub
    End Class
End Namespace
