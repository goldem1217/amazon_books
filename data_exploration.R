library(dplyr)
library('rio')
library('openxlsx')

##### set wd, import csv, add calculated fields #####

wd <- "C:/Users/eg127/Documents/github_repos"
setwd(wd)

# year column removed because all other data fields appear to be aggregates 
# for the entire 2009 - 2019 date range

# filtered out items with price <= 0

df <- import("input/bestsellers with categories.csv") %>%
  select(-Year) %>%
  distinct() %>%
  rename(Rating = "User Rating") %>%
  mutate(Popularity = Reviews*Rating) %>%
  mutate(Sales = Reviews*Price) %>%
  filter(Price > 0) %>%
  arrange(Name)

nonfic <- df %>%
  filter(Genre == "Non Fiction") 

fic <- df %>%
  filter(Genre == "Fiction")

nf_avg_profit <- round(sum(c(nonfic$Sales))/nrow(nonfic))
f_avg_profit <- round(sum(c(fic$Sales))/nrow(fic))

nf_avg_popularity <- round(sum(c(nonfic$Popularity))/nrow(nonfic))
f_avg_popularity <- round(sum(c(fic$Popularity))/nrow(fic))

# Fiction has fewer titles than nonfiction, and on average fiction
# books are more popular (number of reviews * average rating)
# and more profitable (number of reviews * list price) than nonfiction

##### most popular books #####
df1 <- fic %>%
  arrange(desc(Popularity)) %>%
  mutate(toplist = "popularity - fiction") %>%
  top_n(10, Popularity)

df2 <- nonfic %>%
  arrange(desc(Popularity)) %>%
  mutate(toplist = "popularity - nonfiction") %>%
  top_n(10, Popularity)

most_popular <- df1 %>%
  full_join(df2)

##### most profitable books #####
df1 <- fic %>%
  arrange(desc(Sales)) %>%
  mutate(toplist = "profit - fiction") %>%
  top_n(10, Sales)

df2 <- nonfic %>%
  arrange(desc(Sales)) %>%
  mutate(toplist = "profit - nonfiction") %>%
  top_n(10, Sales)

most_profit <- df1 %>%
  full_join(df2)

##### bestselling authors #####
bestsellers <- df %>%
  group_by(Genre, Author) %>% 
  mutate(AuthorProfit = sum(Sales)) %>%
  mutate(AuthorPopularity = sum(Popularity))%>%
  select(c(Author,
           Genre,
           AuthorProfit,
           AuthorPopularity)) %>%
  ungroup() %>%
  distinct() %>%
  arrange(Author)

df1 <- bestsellers %>%
  filter(Genre == "Fiction") %>%
  mutate(toplist = "profit - author") %>%
  top_n(10, AuthorProfit)

df1[nrow(df1)+1,] <- NA

df2 <- bestsellers %>%
  filter(Genre == "Non Fiction") %>%
  mutate(toplist = "popularity - author") %>%
  top_n(10, AuthorPopularity)
  
##### toplists #####

book_lists <- most_popular %>%
  full_join(most_profit) %>%
  add_row(Name = NA,
          Author = NA,
          Rating = NA,
          Reviews = NA,
          Price = NA,
          Genre = NA,
          Popularity = NA,
          Sales = NA,
          toplist = NA,
          .before = 11) %>%
  add_row(Name = NA,
          Author = NA,
          Rating = NA,
          Reviews = NA,
          Price = NA,
          Genre = NA,
          Popularity = NA,
          Sales = NA,
          toplist = NA,
          .before = 22) %>%
  add_row(Name = NA,
          Author = NA,
          Rating = NA,
          Reviews = NA,
          Price = NA,
          Genre = NA,
          Popularity = NA,
          Sales = NA,
          toplist = NA,
          .before = 33) %>%
  rename(Toplist = toplist)

author_lists <- df1 %>%
  full_join(df2) %>%
  rename(c("Author Profit" = AuthorProfit,
           "Author Popularity" = AuthorPopularity,
           Toplist = toplist))

##### export formatted .xlsx #####

file_name <- "output/2009-2019 Book Report.xlsx"

wb <- createWorkbook()

header <- createStyle(
  fgFill = "goldenrod1",
  halign = "center",
  valign = "top",
  textDecoration = "bold",
  wrapText = TRUE,
)

sheet1 <- addWorksheet(wb, "Amazon Books")
sheet2 <- addWorksheet(wb, "Book Toplists")
sheet3 <- addWorksheet(wb, "Amazon Authors")
sheet4 <- addWorksheet(wb, "Author Toplists")

writeData(wb, sheet1, df, headerStyle = header,
          withFilter = TRUE)
writeData(wb, sheet2, book_lists, headerStyle = header)
writeData(wb, sheet3, bestsellers, headerStyle = header,
          withFilter = TRUE)
writeData(wb, sheet4, author_lists, headerStyle = header)

saveWorkbook(wb, file_name, overwrite = TRUE)