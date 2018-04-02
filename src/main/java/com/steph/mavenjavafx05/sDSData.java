/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.steph.mavenjavafx05;

import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;

/**
 *
 * @author stephane
 */
public class sDSData {
    private String      m_Nom;
    private LocalDate   m_Date;
    private String      m_Debut;
    private String      m_Duree;
    private String      m_DeltaTiersTemps;
    private String      m_Salle;
    private boolean     m_AvecTiersTemps;
        
    public sDSData() {
        m_AvecTiersTemps=false;
    }
    public void setNom(String nom) {
        m_AvecTiersTemps=false; // Changement de DS, remise à zero des tiers temps présents
        m_Nom = nom.toUpperCase();
    }
    public String getNom() {
        return m_Nom;
    }
    public void setDate(LocalDate localDate) {
        m_Date=localDate;
    }
    public String getDate() {
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("EEEE dd MMMM YYYY");
        return formatter.format(m_Date);             
    }
    public void setDebut(String debut) {
        m_Debut = debut;
    }
    public String getDebut() {
        return m_Debut;
    }
    public void setDuree(String duree) {
        m_Duree = duree;
    }
    public String getDuree() {
        return m_Duree;
    }
    public void setDeltaTT(String delta) {
        m_DeltaTiersTemps = delta;
    }
    public String getDeltaTT() {
        return m_DeltaTiersTemps;
    }
    public void setSalle(String salle) {
        m_Salle=salle;
    }
    public String getSalle() {
        return m_Salle;
    }
    public String getFin() {
        String[] parts = m_Debut.split("h");
        int heureDebut = Integer.parseInt(parts[0]);
        int minutesDebut = Integer.parseInt(parts[1]);
        parts = m_Duree.split("h");
        int heureDuree = Integer.parseInt(parts[0]);
        int minutesDuree = Integer.parseInt(parts[1]);
        int minutesFin = minutesDebut + minutesDuree + 60*heureDebut+60*heureDuree;
        int heureFin = (minutesFin/60);
        minutesFin -= 60*heureFin;
        if (minutesFin==0) {
            return (heureFin+"h00");
        }
        return (heureFin+"h"+minutesFin);
    }

    public String getDebutTT() {
        if (!m_DeltaTiersTemps.contains("-")) {
            return getDebut();
        }

        String[] parts = m_Debut.split("h");
        int heureDebut = Integer.parseInt(parts[0]);
        int minutesDebut = Integer.parseInt(parts[1]);
        parts = m_DeltaTiersTemps.split("h");
        int heureDeltaTT = Integer.parseInt(parts[0]);
        int minutesDeltaTT = Integer.parseInt(parts[1]);

        int minutesDebutTT = heureDebut*60 + minutesDebut - Math.abs(heureDeltaTT)*60 - minutesDeltaTT;
        int heureDebutTT = (minutesDebutTT/60);
        minutesDebutTT -= 60*heureDebutTT;
        if (minutesDebutTT==0) {
            return (heureDebutTT+"h00");                
        }
        return (heureDebutTT+"h"+minutesDebutTT);
    }

    public String getFinTT() {
        String[] parts = m_Debut.split("h");
        int heureDebut = Integer.parseInt(parts[0]);
        int minutesDebut = Integer.parseInt(parts[1]);
        parts = m_Duree.split("h");
        int heureDuree = Integer.parseInt(parts[0]);
        int minutesDuree = Integer.parseInt(parts[1]);
        int minutesFin = minutesDebut + minutesDuree + 60*heureDebut+60*heureDuree;

        parts = m_DeltaTiersTemps.split("h");
        int heureDeltaTT = Integer.parseInt(parts[0]);
        int minutesDeltaTT = Integer.parseInt(parts[1]);

        if (!m_DeltaTiersTemps.contains("-")) {
            minutesFin += heureDeltaTT*60 + minutesDeltaTT;
        }
        int heureFin = (minutesFin/60);
        minutesFin -= 60*heureFin;
        if (minutesFin==0) {
            return (heureFin+"h00");
        }
        return (heureFin+"h"+minutesFin);
    }

    public void setTiersTemps(boolean tierstemps) {
        m_AvecTiersTemps=tierstemps;
    }

    public boolean avecTiersTemps() {
        return m_AvecTiersTemps;
    }




    public static void main(String[] args)  {

        sDSData data = new sDSData();

        data.setNom("ds if");
        LocalDate localDate = LocalDate.now();
        data.setDate(localDate);
        System.out.println(data.getDate());
        System.out.println("---------------------");
        data.setDebut("10h45");
        data.setDuree("1h45");
        data.setDeltaTT("-1h30");

        System.out.println("Début : "+data.getDebut()+"   Durée : "+data.getDuree()+" DeltaTT : "+data.getDeltaTT());
        System.out.println("-> Début   : "+data.getDebut());
        System.out.println("-> Fin     : "+data.getFin());
        System.out.println("-> DébutTT : "+data.getDebutTT());
        System.out.println("-> FinTT   : "+data.getFinTT());
    }             
}
