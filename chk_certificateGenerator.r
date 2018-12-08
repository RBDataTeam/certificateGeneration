setwd("F:/Reap_Benefit/Data/Certificate")
oldw <- getOption("warn")
options(warn = -1)

suppressMessages(library(officer))
suppressMessages(library(ggplot2))
suppressMessages(library(tidyverse))
suppressMessages(library(readxl))

#Image Variables for Text
img_src <- c(
  "Action Ant" = c("pic" = "const/ActionAnt.png","height" = 1.61,"width"=1.13,
                   "left" = 5.58, "top" = 7.39),
  "Builder Bear" = c("pic" = "const/BuilderBear.png", "height" = 1.80,"width"=1.03,
                     "left" = 5.63, "top" = 7.39),
  "Campaign Chameleon" = c("pic" = "const/CampaignChameleon.png", "height" = 1.46,"width"=1.63,
                           "left" = 5.33, "top" = 7.39),
  "Curious Cat" =  c("pic" = "const/CuriousCat.png", "height" = 1.60,"width"=1.03,
                     "left" = 5.61, "top" = 7.39),
  "Hands On Hippo" = c("pic" = "const/HandsonHippo.png", "height" = 1.70,"width"=1.35,
                       "left" = 5.47, "top" = 7.39 ),
  "Reporting Rhino" = c("pic" = "const/ReportingRhino.png", "height" = 1.70,"width"=1.18,
                        "left" = 5.56, "top" = 7.39),
  "Techno Tiger" = c("pic" = "const/TechnoTiger.png","height" = 1.61,"width"=1.38,
                     "left" = 5.46, "top" = 7.39)
)

# Graph Positions of Skills
position <- c("Data Orientation", "Hands-on Skills",
              "Citizenship", "Critical Thinking", "Problem Solving", 
              "Communication", "Community Collaboration",
              "Grit", "Applied Empathy", "Entrepreneurship")

# Chart Colors
literacyColor <- "#86D1F2"
proficiencyColor <- "#EFC473"
characterColor <- "#EE8962"


#Read Data
trainxl <- readxl::read_xlsx("Data/data.xlsx") #raw data
skilldf <- read.csv("const/Skills - new.csv") # Skill names

# Calculate Levels based on L0 - L7
trainxl$Level_Certi <- ifelse(trainxl$`Level Group` == "L0", "Level 0",
                              ifelse(trainxl$`Level Group` == "L1",  "Level 1",
                                     ifelse(trainxl$`Level Group` == "L2",  "Level 2",
                                            ifelse(trainxl$`Level Group` == "L3",  "Level 3",
                                                   ifelse(trainxl$`Level Group` == "L4",  "Level 4",
                                                          ifelse(trainxl$`Level Group` == "L5",  "Level 5",
                                                                 ifelse(trainxl$`Level Group` == "L6",  "Level 6", 
                                                                        ifelse(trainxl$`Level Group` == "L7",  "Level 7","Error"
                                                                        ))))))))


# to find out how many certificates need to be created
certificateNum <- dim(trainxl)[1]

# What is the certificate template
certificate <- read_pptx("const/certificate_template2.pptx")

i = 1
for(i in 1:certificateNum){
  df1 <- trainxl[i,]
  
  
  studentName <- df1$Student_Name
  studentSNI <- df1$`Sum of Index`
  studentPersona <- df1$Persona
  studentLevel <- df1$Level_Certi
  studenNum <- df1$SLNUM
  
  
  
  df1 <- gather(df1,key = "Details",value = "Actions",
                -c(1:7, 18:30 ))
  
  df1 <- merge(df1, skilldf, by.x = "Details", by.y = "index")
  
  
  
  # We find max number of actions for dividing into grid lines
  gridStop <- max(df1$Actions)
  
  # To keep margin we multiply by 1.2
  gridBar <- gridStop * 1.2
  
  df1$skills <- factor(df1$skills, levels = position) 
  
  displaydf <- df1[, c(8,23)] #Keep only Student Name and Skills
  
  # The function below creates gridlines by dividing the max
  # Number of Actions into 15 parts
  numGridlines <- 15 # 15 Grid lines
  for(i in 1:numGridlines){
    displaydf[[paste0("Action_",i)]] <- gridBar - ((i * gridBar)/numGridlines)
  }
  
  displaydf <- gather(displaydf, key= "More_Details",value = "Actions", 
                      -c(1:2))
  
  # Creating Plot
  p <- ggplot(df1,mapping = aes(skills, Actions, fill = skills)) +
    geom_bar(width = 1.00, stat = "identity", alpha = 0.95) +
    geom_bar(data = displaydf, width = 1.00, fill = "grey", 
             stat = "identity", alpha = 0, color = "#bad8e1", 
             position = "dodge")+
    theme_bw() +
    theme(axis.ticks = element_blank(),
          axis.title = element_blank(),
          axis.line = element_blank(),
          axis.text.y = element_blank(),
          axis.text.x = element_blank(),
          # axis.text.x = element_text(colour = "black",
          #                            hjust = 0,
          #                            vjust = 0,
          #                            family = "Exo 2",
          #                            size = 15),
          panel.border = element_blank(),
          legend.position = "none",
          panel.grid.major = element_blank(),
          panel.grid.major.x = element_blank()
    ) +
    scale_x_discrete(limits = position)+
    scale_fill_manual(values = c(
      "Data Orientation" = literacyColor,
      "Hands-on Skills" = literacyColor,
      "Citizenship" = literacyColor,
      "Critical Thinking" = proficiencyColor,
      "Problem Solving" = proficiencyColor,
      "Communication" = proficiencyColor,
      "Community Collaboration"= characterColor,
      "Grit" = characterColor,
      "Applied Empathy" = characterColor,
      "Entrepreneurship" = characterColor
    )) + coord_polar(theta = "x", start = 0, direction = 1)
  
  # Inserting Plot Annotations
  p <- p + annotate("text", 
                    x = 1.2, y = gridBar+2 , 
                    label = "Data \nOrientation", 
                    family = "Exo 2",
                    size = 2.5)
  p <- p + annotate("text", 
                    x = 2.1, y = gridBar+2, 
                    label = "Hands On \nSkills", 
                    family = "Exo 2",
                    size = 2.5)
  p <- p + annotate("text", 
                    x = 3, y = gridBar+3, 
                    label = "Citizenship", 
                    family = "Exo 2",
                    size = 2.5)
  p <- p + annotate("text", 
                    x = 3.9, y = gridBar+2, 
                    label = "Critical \n Thinking", 
                    family = "Exo 2",
                    size = 2.5)
  p <- p + annotate("text", 
                    x = 4.8, y = gridBar+2,
                    label = "Problem \n Solving",
                    family = "Exo 2",
                    size = 2.5)
  p <- p + annotate("text", 
                    x = 6.2, y = gridBar+1.7, 
                    label = "Communication", 
                    family = "Exo 2",
                    size = 2.5)
  p <- p + annotate("text", 
                    x = 7.1, y = gridBar+3,
                    label = "Community \nCollaboration",
                    family = "Exo 2",
                    size = 2.5)
  p <- p + annotate("text", 
                    x = 8, y = gridBar+1.7,
                    label = "Grit",
                    family = "Exo 2",
                    size = 2.5)
  p <- p + annotate("text", 
                    x = 8.9, y = gridBar+2.5,
                    label = "Applied \nEmpathy",
                    family = "Exo 2",
                    size = 2.5)
  p <- p + annotate("text", 
                    x = 9.8, y = gridBar+2,
                    label = "Entrepreneurship",
                    family = "Exo 2",
                    size = 2.5)
  
# Adding student details on certificate  
  certificate <- certificate %>% 
    add_slide(layout = "Title and Content", master = "Office Theme")%>% 
    ph_empty_at(left=1.54, top=1.88,
                width=4.43, height=1.04)%>%
    ph_empty_at(left=5.00, top=9.28,
                width=2.5, height=0.47)%>%
    ph_add_par(level = 1, id_chr = 2) %>%
    ph_add_text(paste0("Congratulations ", studentName, "!"),
                type = "body",
                id_chr = 2,
                style = fp_text(color = "white", font.size = 28)) %>%
    ph_add_par(level = 1, id_chr = 3) %>%
    ph_add_text(paste0(studentName, " is a ", studentLevel," ", studentPersona),
                type = "body",
                id_chr = 3,
                style = fp_text(color = "black", font.size = 10))
  
  
  # Inserting students Skill Plot
  
  certificate <- ph_with_gg_at(x = certificate, value = p,
                               left = 1.739726, top = 6.868362,
                               height = 2.87, width = 3.3)
  
  # Inserting students Solve Ninja Image
  certificate <- ph_with_img_at(x = certificate,
                                src = img_src[paste0(studentPersona,".pic")],
                                left = as.numeric(img_src[paste0(studentPersona,".left")]), 
                                top = as.numeric(img_src[paste0(studentPersona,".top")]),
                                height = as.numeric(img_src[paste0(studentPersona,".height")]), 
                                width = as.numeric(img_src[paste0(studentPersona,".width")]))
  
  
}
#Print the final certificate
print(certificate, paste0("output/certificates.pptx"))

#options(warn = oldw)
